# api_clients/sfera_api_client.py
import requests
import pandas as pd
import logging
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional, Any
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from requests.exceptions import RequestException, Timeout, ConnectionError

from config import SFERA_RETRY_ATTEMPTS, SFERA_API_TIMEOUT

logger = logging.getLogger('sfera_api')


class SferaAPIClient:
    """
    Клиент для работы с API Сфера Документы (Сфера Курьер).
    Позволяет получить список сотрудников организации.
    """
    def __init__(self, base_url: str, username: str, password: str, timeout: int = 60):
        self.base_url = base_url.rstrip('/')
        self.username = username
        self.password = password
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({
            'Content-Type': 'application/xml',
            'Accept': 'application/xml'
        })
        self._token = None
        self._admin_role_id = None

    def _logon(self) -> str:
        """
        POST /api/auth/logon
        Выполняет аутентификацию и получает токен.
        """
        url = f"{self.base_url}/api/auth/logon"
        # Формируем XML тело запроса
        xml_body = f"""<?xml version="1.0" encoding="UTF-8"?>
<Credentials>
    <Username>{self.username}</Username>
    <Password>{self.password}</Password>
</Credentials>"""
        logger.debug("Выполнение аутентификации в Сфера API")
        response = self.session.post(url, data=xml_body, timeout=self.timeout)
        if response.status_code != 200:
            raise RuntimeError(f"Ошибка аутентификации: {response.status_code} - {response.text}")
        # Парсим XML ответ: ожидается <LogonResponse><Token>...</Token></LogonResponse>
        try:
            root = ET.fromstring(response.content)
            token_elem = root.find('.//Token')
            if token_elem is None or not token_elem.text:
                raise ValueError("Не удалось извлечь токен из ответа")
            token = token_elem.text.strip()
            logger.info("Аутентификация успешна, токен получен")
            return token
        except ET.ParseError as e:
            raise RuntimeError(f"Ошибка парсинга XML при аутентификации: {e}")

    def _ensure_token(self):
        """Проверяет наличие токена, при необходимости выполняет вход."""
        if not self._token:
            self._token = self._logon()
            # Устанавливаем заголовок авторизации для всех последующих запросов
            self.session.headers.update({'Auth-Token': self._token})

    @retry(
        stop=stop_after_attempt(SFERA_RETRY_ATTEMPTS),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((ConnectionError, Timeout))
    )
    def _post_xml(self, endpoint: str, data: Optional[str] = None) -> ET.Element:
        """
        Выполняет POST-запрос с XML телом, возвращает корневой элемент ответа.
        """
        self._ensure_token()
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        response = self.session.post(url, data=data, timeout=self.timeout)
        if response.status_code == 401:
            # Возможно, токен истёк – сбрасываем и пробуем ещё раз
            self._token = None
            self.session.headers.pop('Auth-Token', None)
            self._ensure_token()
            response = self.session.post(url, data=data, timeout=self.timeout)
        if response.status_code != 200:
            raise RuntimeError(f"Ошибка запроса {endpoint}: {response.status_code} - {response.text}")
        try:
            return ET.fromstring(response.content)
        except ET.ParseError as e:
            raise RuntimeError(f"Ошибка парсинга XML ответа от {endpoint}: {e}")

    def get_role_list(self) -> List[Dict[str, Any]]:
        """
        POST /api/helper/roleList
        Возвращает список всех ролей, доступных в системе.
        """
        logger.debug("Запрос списка ролей")
        root = self._post_xml("api/helper/roleList")
        roles = []
        # Ожидаем структуру: <ArrayOfRole> <Role>...</Role> ... </ArrayOfRole>
        for role_elem in root.findall('.//Role'):
            role_id_elem = role_elem.find('Id')
            name_elem = role_elem.find('Name')
            role = {
                'id': int(role_id_elem.text) if role_id_elem is not None and role_id_elem.text else None,
                'name': name_elem.text if name_elem is not None else ''
            }
            if role['id'] is not None:
                roles.append(role)
        logger.info(f"Получено {len(roles)} ролей")
        return roles

    def _get_admin_role_id(self) -> int:
        """
        Определяет ID роли 'Администратор системы'.
        Если роль не найдена, возвращает 2 (типовое значение).
        """
        if self._admin_role_id is not None:
            return self._admin_role_id
        roles = self.get_role_list()
        # Ищем роль с названием "Администратор системы" (можно уточнить)
        admin_role = next((r for r in roles if r['name'] and 'администратор' in r['name'].lower()), None)
        if admin_role:
            self._admin_role_id = admin_role['id']
            logger.info(f"ID роли администратора: {self._admin_role_id}")
        else:
            logger.warning("Роль 'Администратор системы' не найдена. Используем fallback ID = 2 (типовое значение)")
            self._admin_role_id = 2
        return self._admin_role_id

    def get_user_list(self) -> List[Dict[str, Any]]:
        """
        POST /api/helper/userList
        Возвращает список пользователей компании.
        """
        logger.debug("Запрос списка пользователей")
        root = self._post_xml("api/helper/userList")
        users = []
        # Ответ: <ArrayOfUserDetails> <UserDetails>...</UserDetails> ... </ArrayOfUserDetails>
        for user_elem in root.findall('.//UserDetails'):
            user = self._parse_user_details(user_elem)
            if user:
                users.append(user)
        logger.info(f"Получено {len(users)} пользователей")
        return users

    def _parse_user_details(self, user_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """Парсит один элемент UserDetails в словарь."""
        try:
            # ФИО
            person_elem = user_elem.find('Person')
            last_name = person_elem.findtext('LastName', '') if person_elem is not None else ''
            first_name = person_elem.findtext('FirstName', '') if person_elem is not None else ''
            middle_name = person_elem.findtext('MiddleName', '') if person_elem is not None else ''
            full_name = ' '.join(filter(None, [last_name, first_name, middle_name])).strip()

            # Активен
            is_active_elem = user_elem.find('IsActive')
            is_active = is_active_elem.text.lower() == 'true' if is_active_elem is not None else False

            # Роли
            roles = []
            roles_elem = user_elem.find('Roles')
            if roles_elem is not None:
                for role_elem in roles_elem.findall('Role'):
                    role_id_elem = role_elem.find('Id')
                    if role_id_elem is not None and role_id_elem.text:
                        roles.append(int(role_id_elem.text))

            # Остальные поля (могут пригодиться)
            login = user_elem.findtext('Login', '')
            email = user_elem.findtext('Email', '')
            company = user_elem.findtext('Company', '')

            return {
                'full_name': full_name,
                'is_active': is_active,
                'roles': roles,
                'login': login,
                'email': email,
                'company': company
            }
        except Exception as e:
            logger.warning(f"Ошибка парсинга элемента UserDetails: {e}")
            return None

    def userlist_to_dataframe(self, users_raw: List[Dict]) -> pd.DataFrame:
        """
        Преобразует список пользователей из API в pandas DataFrame,
        полностью совместимый с загружаемым из Excel.
        Колонки:
            - Сфера_Курьер_ФИО
            - Сфера_Курьер_Активен (Да/Нет)
            - Сфера_Курьер_Администратор (Да/Нет)
        """
        admin_role_id = self._get_admin_role_id()
        data = []
        for user in users_raw:
            if not user['full_name']:
                continue
            is_admin = admin_role_id in user['roles']
            data.append({
                'Сфера_Курьер_ФИО': user['full_name'],
                'Сфера_Курьер_Активен': 'Да' if user['is_active'] else 'Нет',
                'Сфера_Курьер_Администратор': 'Да' if is_admin else 'Нет'
            })
        df = pd.DataFrame(data)
        logger.info(f"Сформирован DataFrame: {len(df)} строк")
        return df


# --- Синглтон для получения клиента ---
_sfera_client = None

def get_sfera_client():
    """Возвращает экземпляр клиента Сфера API (синглтон)."""
    global _sfera_client
    if _sfera_client is None:
        from config import SFERA_API_BASE_URL, SFERA_USERNAME, SFERA_PASSWORD, SFERA_API_TIMEOUT
        if not SFERA_USERNAME or not SFERA_PASSWORD:
            logger.error("SFERA_USERNAME или SFERA_PASSWORD не заданы. Невозможно создать клиент API.")
            return None
        _sfera_client = SferaAPIClient(
            base_url=SFERA_API_BASE_URL,
            username=SFERA_USERNAME,
            password=SFERA_PASSWORD,
            timeout=SFERA_API_TIMEOUT
        )
    return _sfera_client
