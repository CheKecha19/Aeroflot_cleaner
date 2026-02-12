# api_clients/diadoc_api_client.py
import requests
import pandas as pd
import logging
from typing import List, Dict, Optional, Any
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from requests.exceptions import RequestException, Timeout, ConnectionError

from config import DIADOC_PAGE_SIZE, DIADOC_RETRY_ATTEMPTS

logger = logging.getLogger('diadoc_api')


class DiadocAPIClient:
    """
    Клиент для работы с API Контур.Диадок.
    Позволяет получить список сотрудников организации.
    """
    def __init__(self, api_key: str, base_url: str, box_id: Optional[str] = None, timeout: int = 60):
        self.api_key = api_key
        self.base_url = base_url.rstrip('/')
        self.box_id = box_id
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({
            'Authorization': f'Bearer {self.api_key}',
            'Accept': 'application/json; charset=utf-8',
            'Content-Type': 'application/json'
        })
        # Кэш для идентификатора ящика, чтобы не запрашивать каждый раз
        self._resolved_box_id = None

    def _get(self, endpoint: str, params: Optional[Dict] = None) -> Dict[str, Any]:
        """Базовый GET-запрос с повторными попытками и обработкой ошибок."""
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        try:
            response = self.session.get(url, params=params, timeout=self.timeout)
            # Пробрасываем исключения для специфических кодов
            if response.status_code == 401:
                raise PermissionError("401 Unauthorized: неверный или отсутствует API-ключ")
            if response.status_code == 402:
                raise RuntimeError("402 Payment Required: закончилась подписка на API Диадок")
            if response.status_code == 403:
                raise PermissionError("403 Forbidden: доступ запрещён. Убедитесь, что ключ принадлежит администратору организации")
            if response.status_code == 404:
                raise ValueError("404 Not Found: ящик с указанным boxId не найден")
            if response.status_code == 409:
                raise RuntimeError("409 Conflict: превышен лимит запросов в день")
            response.raise_for_status()
            return response.json()
        except (ConnectionError, Timeout) as e:
            logger.error(f"Сетевая ошибка при запросе к {url}: {e}")
            raise
        except requests.HTTPError as e:
            logger.error(f"HTTP ошибка {response.status_code}: {response.text}")
            raise

    @retry(
        stop=stop_after_attempt(DIADOC_RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((ConnectionError, Timeout))
    )
    def get_my_organizations(self) -> List[Dict]:
        """
        GET /GetMyOrganizations
        Возвращает список организаций, доступных текущему пользователю.
        """
        logger.debug("Запрос /GetMyOrganizations")
        data = self._get("GetMyOrganizations", params={"autoRegister": "false"})
        orgs = data.get("Organizations", [])
        logger.info(f"Получено организаций: {len(orgs)}")
        return orgs

    def _resolve_box_id(self, preferred_box_id: Optional[str] = None) -> str:
        """
        Определяет BoxId для дальнейших запросов.
        Если передан preferred_box_id, использует его.
        Иначе запрашивает список организаций и берёт BoxId первой организации.
        Результат кэшируется.
        """
        if self._resolved_box_id:
            return self._resolved_box_id

        if preferred_box_id:
            self._resolved_box_id = preferred_box_id
            logger.info(f"Используется явно указанный BoxId: {preferred_box_id}")
            return preferred_box_id

        # Получаем организации и берём первый доступный ящик
        orgs = self.get_my_organizations()
        if not orgs:
            raise RuntimeError("Не найдено ни одной организации, доступной по данному API-ключу")
        # Ищем первый ящик в первой организации
        first_org = orgs[0]
        boxes = first_org.get("Boxes", [])
        if not boxes:
            raise RuntimeError(f"В организации {first_org.get('OrgId')} нет ни одного ящика")
        box_id = boxes[0].get("BoxId")
        if not box_id:
            raise RuntimeError("Не удалось получить BoxId из структуры организации")
        self._resolved_box_id = box_id
        logger.info(f"Автоматически выбран BoxId: {box_id} (организация: {first_org.get('ShortName', first_org.get('FullName', 'Unknown'))})")
        return box_id

    @retry(
        stop=stop_after_attempt(DIADOC_RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((ConnectionError, Timeout))
    )
    def get_employees_page(self, box_id: str, page: int = 1, count: int = 50) -> Dict[str, Any]:
        """
        GET /GetEmployees
        Возвращает одну страницу сотрудников.
        """
        params = {
            "boxId": box_id,
            "page": page,
            "count": min(count, 50)  # ограничиваем максимумом API
        }
        logger.debug(f"Запрос /GetEmployees: boxId={box_id}, page={page}, count={params['count']}")
        return self._get("GetEmployees", params=params)

    def get_all_employees(self, box_id: Optional[str] = None) -> List[Dict]:
        """
        Получить ВСЕХ сотрудников организации с пагинацией.
        Возвращает список сырых объектов Employee.
        """
        box_id = self._resolve_box_id(box_id or self.box_id)
        all_employees = []
        page = 1
        total_count = None

        logger.info(f"Начинаем сбор сотрудников для boxId: {box_id}")
        while True:
            data = self.get_employees_page(box_id, page=page, count=DIADOC_PAGE_SIZE)
            employees = data.get("Employees", [])
            all_employees.extend(employees)

            # При первом запросе получаем общее количество
            if total_count is None:
                total_count = data.get("TotalCount", 0)
                logger.info(f"Всего сотрудников по данным API: {total_count}")

            logger.debug(f"Страница {page}: получено {len(employees)} сотрудников, всего собрано {len(all_employees)}")

            # Условие завершения: если получили меньше, чем запрашивали, или собрали уже всех
            if len(employees) < DIADOC_PAGE_SIZE or (total_count and len(all_employees) >= total_count):
                break
            page += 1

        logger.info(f"Сбор сотрудников завершён: всего {len(all_employees)} записей")
        return all_employees

    @staticmethod
    def _extract_full_name(fullname_dict: Dict[str, str]) -> str:
        """Собирает ФИО из структуры FullName."""
        parts = [
            fullname_dict.get('LastName', ''),
            fullname_dict.get('FirstName', ''),
            fullname_dict.get('MiddleName', '')
        ]
        return ' '.join(filter(None, parts)).strip()

    def employees_to_dataframe(self, employees_raw: List[Dict]) -> pd.DataFrame:
        """
        Преобразует список сотрудников (из API) в pandas DataFrame,
        полностью совместимый с тем, который загружается из Excel.
        Колонки:
            - Контур_Диадок_ФИО
            - Контур_Диадок_Администратор (да/нет)
            - Контур_Диадок_статус (активна/заблокирована)
        """
        data = []
        skipped = 0
        for emp in employees_raw:
            try:
                # Извлекаем ФИО
                user = emp.get('User', {})
                fullname_dict = user.get('FullName', {})
                fio = self._extract_full_name(fullname_dict)
                if not fio:
                    skipped += 1
                    continue

                # Права и статус
                permissions = emp.get('Permissions', {})
                is_admin = permissions.get('IsAdministrator', False)
                auth_perm = permissions.get('AuthorizationPermission', {})
                is_blocked = auth_perm.get('IsBlocked', False)

                data.append({
                    'Контур_Диадок_ФИО': fio,
                    'Контур_Диадок_Администратор': 'да' if is_admin else 'нет',
                    'Контур_Диадок_статус': 'заблокирована' if is_blocked else 'активна'
                })
            except Exception as e:
                logger.warning(f"Ошибка обработки сотрудника: {e}. Запись пропущена.")
                skipped += 1
                continue

        if skipped:
            logger.warning(f"Пропущено {skipped} записей из-за отсутствия ФИО или ошибок структуры")
        df = pd.DataFrame(data)
        logger.info(f"Сформирован DataFrame: {len(df)} строк")
        return df


# --- Синглтон для получения клиента ---
_diadoc_client = None

def get_diadoc_client():
    """Возвращает экземпляр клиента (синглтон)."""
    global _diadoc_client
    if _diadoc_client is None:
        from config import DIADOC_API_KEY, DIADOC_API_BASE_URL, DIADOC_BOX_ID, DIADOC_API_TIMEOUT
        if not DIADOC_API_KEY:
            logger.error("DIADOC_API_KEY не задан. Невозможно создать клиент API.")
            return None
        _diadoc_client = DiadocAPIClient(
            api_key=DIADOC_API_KEY,
            base_url=DIADOC_API_BASE_URL,
            box_id=DIADOC_BOX_ID or None,
            timeout=DIADOC_API_TIMEOUT
        )
    return _diadoc_client
