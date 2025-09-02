import sys
import json
import time
import os
import webbrowser
from datetime import datetime
from typing import List, Dict, Optional
import requests
from bs4 import BeautifulSoup
import re
import smtplib
import socket
import dns.resolver
import warnings
import urllib3
from urllib3.exceptions import InsecureRequestWarning
urllib3.disable_warnings(InsecureRequestWarning)
warnings.filterwarnings('ignore', message='Unverified HTTPS request')
import smtplib
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QComboBox, QSpinBox, QCheckBox, QGroupBox, QFileDialog,
    QMessageBox, QHeaderView, QAbstractItemView, QProgressBar,
    QMenu, QAction, QTextEdit, QDialog, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon, QCursor

import pandas as pd
import requests


class CategoryEditDialog(QDialog):
    def __init__(self, categories, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Редактирование категорий поиска")
        self.setModal(True)
        self.resize(400, 300)
        layout = QVBoxLayout(self)
        label = QLabel("Категории для поиска (одна на строку):")
        layout.addWidget(label)
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText('\n'.join(categories))
        layout.addWidget(self.text_edit)
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal, self
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
    def get_categories(self):
        text = self.text_edit.toPlainText()
        categories = [line.strip() for line in text.split('\n') if line.strip()]
        return categories

class EmailValidator:
    def __init__(self):
        self.common_prefixes = ['info', 'contact']
        self.session = requests.Session()
        self.session.verify = False
        self.mx_cache = {}
        
        self.domain_cache = {}
        
        self.contact_pages = [
            '/kontaktnaya-informatsiya/',
            '/kontakty/',
            '/contacts/',
            '/contact/',
            '/contact-us/',
            '/about/contacts/',
            '/o-nas/kontakty/',
            '/kontakt/',
            '/contact.html',
            '/contacts.html',
            '/contact.php',
            '/about/',
            '/o-nas/',
            '/info/'
        ]
        
        self.domain_replacements = {
            'med-b-': '',
            '.gosweb.gosuslugi.ru': '.gosuslugi.ru',
            '-r73.': '73.',
        }

    def get_domain_key(self, website):
        if not website:
            return None
        
        domain = self.extract_base_domain(website)
        if not domain:
            return None
        
        domain = domain.lower()
        for old, new in self.domain_replacements.items():
            domain = domain.replace(old, new)
        
        return domain

    def find_emails_for_website(self, website, smart_mode=True):
        domain_key = self.get_domain_key(website)
        
        if domain_key and domain_key in self.domain_cache:
            cached_emails = self.domain_cache[domain_key]
            return {'parsed': cached_emails, 'verified': [], 'suggested': []}
        
        print(f"Парсинг домена: {domain_key}")
        if smart_mode:
            parsed = self.parse_emails_from_website_smart(website)
        else:
            parsed = self.parse_single_page(website, timeout=3)
        
        if domain_key:
            self.domain_cache[domain_key] = parsed
            if parsed:
                print(f"Найдены email на {domain_key}: {', '.join(parsed)}")
        
        return {'parsed': parsed, 'verified': [], 'suggested': []}

    def extract_base_domain(self, website):
        if not website:
            return None
            
        if website.startswith(('http://', 'https://')):
            website = website.split('://', 1)[1]
        
        domain_part = website.split('/')[0]
        
        return domain_part

    def fix_government_domain(self, url):
        original_url = url
        
        for old, new in self.domain_replacements.items():
            if old in url:
                url = url.replace(old, new)
        
        return url

    def check_site_availability(self, url, timeout=3):
        try:
            response = self.session.head(url, timeout=timeout, allow_redirects=True)
            return response.status_code < 500
        except:
            return False

    def extract_emails_from_html(self, html_content):
        emails = set()
        
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            
            for link in soup.find_all('a', href=True):
                href = link.get('href', '')
                if href.startswith('mailto:'):
                    email = href.replace('mailto:', '').split('?')[0].strip()
                    if '@' in email and '.' in email:
                        emails.add(email.lower())
            
            text = soup.get_text()
            email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
            found_emails = re.findall(email_pattern, text)
            
            for email in found_emails:
                email = email.lower().strip()
                if '@' in email and '.' in email:
                    emails.add(email)
            
            for tag in soup.find_all(True):
                for attr, value in tag.attrs.items():
                    if isinstance(value, str) and '@' in value:
                        found_emails = re.findall(email_pattern, value)
                        for email in found_emails:
                            emails.add(email.lower().strip())
            
        except Exception as e:
            pass
        
        valid_emails = []
        exclude_patterns = [
            'example.com', 'test.com', 'domain.com', 'email.com',
            'noreply', 'no-reply', 'donotreply', 'webmaster@', 
            'postmaster@', 'abuse@', 'support@example',
            'user@example', 'name@example', '@example',
            '.png', '.jpg', '.gif', '.svg'
        ]
        
        for email in emails:
            if (not any(pattern in email for pattern in exclude_patterns) and
                len(email) > 5 and 
                email.count('@') == 1 and
                not email.startswith('@') and
                not email.endswith('@')):
                valid_emails.append(email)
        
        return list(set(valid_emails))[:3]

    def parse_single_page(self, url, timeout=5):

        try:

            if 'gosuslugi.ru' in url or 'gosweb.' in url:
                url = self.fix_government_domain(url)
                
                if not self.check_site_availability(url, timeout=2):
                    return []
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
                'Accept-Encoding': 'gzip, deflate, br',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1'
            }
            
            if any(gov_domain in url.lower() for gov_domain in 
                   ['gosuslugi.ru', 'gov.ru', 'minzdrav']):
                timeout = 8
            
            response = self.session.get(url, headers=headers, timeout=timeout, allow_redirects=True)
            
            if response.status_code != 200:
                return []
            
            emails = self.extract_emails_from_html(response.text)
            
            return emails
            
        except:
            return []

    def generate_contact_urls(self, website):
        base_domain = self.extract_base_domain(website)
        if not base_domain:
            return []
        
        if 'gosuslugi.ru' in base_domain or 'gosweb.' in base_domain:
            base_domain = self.fix_government_domain(base_domain)
        
        if any(gov_domain in base_domain.lower() for gov_domain in 
               ['gosuslugi.ru', 'gov.ru', 'minzdrav', 'mz']):
            protocol = 'https://'
        elif website.startswith('https://'):
            protocol = 'https://'
        elif website.startswith('http://'):
            protocol = 'http://'
        else:
            protocol = 'https://'
        
        contact_urls = []
        
        if website.startswith(('http://', 'https://')):
            website_without_protocol = website.split('://', 1)[1]
            fixed_url = f"{protocol}{website_without_protocol}"
            fixed_url = self.fix_government_domain(fixed_url)
            contact_urls.append(fixed_url)
        else:
            fixed_url = f"{protocol}{website}"
            fixed_url = self.fix_government_domain(fixed_url)
            contact_urls.append(fixed_url)
        
        main_url = f"{protocol}{base_domain}/"
        if main_url not in contact_urls:
            contact_urls.append(main_url)
        
        pages_to_check = self.contact_pages[:4] if 'gosuslugi.ru' in base_domain else self.contact_pages
        
        for path in pages_to_check:
            contact_url = f"{protocol}{base_domain}{path}"
            if contact_url not in contact_urls:
                contact_urls.append(contact_url)
        
        return contact_urls[:6] 

    def parse_emails_from_website_smart(self, website):
        if not website:
            return []
        
        all_emails = []
        found_emails_set = set()
        contact_urls = self.generate_contact_urls(website)
        
        priority_urls = []
        other_urls = []
        
        for url in contact_urls:
            url_lower = url.lower()
            if any(keyword in url_lower for keyword in 
                   ['kontakt', 'contact', 'about', 'o-nas', 'info']):
                priority_urls.append(url)
            else:
                other_urls.append(url)
        
        for url in priority_urls:
            if len(all_emails) >= 3:
                break
                
            emails = self.parse_single_page(url, timeout=5)
            for email in emails:
                if email not in found_emails_set:
                    found_emails_set.add(email)
                    all_emails.append(email)
        
        if len(all_emails) < 2:
            for url in other_urls:
                if len(all_emails) >= 3:
                    break
                    
                emails = self.parse_single_page(url, timeout=3)
                for email in emails:
                    if email not in found_emails_set:
                        found_emails_set.add(email)
                        all_emails.append(email)
        
        return all_emails[:3] 
class SearchWorker(QThread):
    progress = pyqtSignal(str)
    progress_value = pyqtSignal(int)
    result_found = pyqtSignal(dict)
    search_completed = pyqtSignal()
    error_occurred = pyqtSignal(str)
    
    def __init__(self, api_key: str, city: str, radius: int, categories: List[str]):
        super().__init__()
        self.api_key = api_key
        self.city = city
        self.radius = radius
        self.categories = categories
        self.base_url = "https://catalog.api.2gis.com/3.0"
        self.is_running = True
        
        self.found_ids = set()
        self.found_emails = set()
        self.found_websites = set()
        
        self.parse_websites = False
        self.smart_parsing = True
        
        self.email_validator = EmailValidator()
    
    def run(self):
        try:
            self.progress.emit("Получение координат города...")
            lon, lat = self.get_city_coordinates(self.city)
            
            if not lon or not lat:
                self.error_occurred.emit(f"Не удалось найти город: {self.city}")
                return
            
            self.progress.emit(f"Город найден: {self.city} ({lat:.4f}, {lon:.4f})")
            
            self.progress.emit("Загрузка данных из 2GIS...")
            all_raw_data = []
            
            total_categories = len(self.categories)
            for idx, category in enumerate(self.categories):
                if not self.is_running:
                    break
                
                self.progress.emit(f"Загрузка категории: {category}")
                self.progress_value.emit(int((idx / total_categories) * 50))  # 50% на загрузку
                
                items = self.search_medical_organizations(lon, lat, self.radius, category)
                
                for item in items:
                    if not self.is_running:
                        break
                    
                    contact_info = self.extract_basic_contact_info(item)
                    contact_info['category'] = category
                    
                    if contact_info['id'] and contact_info['id'] not in self.found_ids:
                        all_raw_data.append(contact_info)
                        self.found_ids.add(contact_info['id'])
            
            self.progress.emit(f"Загружено {len(all_raw_data)} учреждений")
            
            if self.parse_websites:
                self.progress.emit("Группировка по доменам...")
                domain_groups = {}
                items_without_websites = []
                
                for item in all_raw_data:
                    website = item.get('website', '').strip()
                    if website:
                        domain = self.email_validator.get_domain_key(website)
                        if domain:
                            if domain not in domain_groups:
                                domain_groups[domain] = {
                                    'website': website,
                                    'items': []
                                }
                            domain_groups[domain]['items'].append(item)
                        else:
                            items_without_websites.append(item)
                    else:
                        items_without_websites.append(item)
                
                self.progress.emit(f"Найдено {len(domain_groups)} уникальных доменов")
                
                parsed_domains = 0
                total_domains = len(domain_groups)
                
                for domain, group_info in domain_groups.items():
                    if not self.is_running:
                        break
                    
                    parsed_domains += 1
                    progress_percent = 50 + int((parsed_domains / total_domains) * 40)  # 40% на парсинг
                    self.progress_value.emit(progress_percent)
                    self.progress.emit(f"Парсинг домена {parsed_domains}/{total_domains}: {domain}")
                    
                    email_results = self.email_validator.find_emails_for_website(
                        group_info['website'], smart_mode=self.smart_parsing
                    )
                    
                    found_emails = email_results.get('parsed', [])
                    
                    for item in group_info['items']:
                        if not item.get('email') and found_emails:  
                            new_emails = []
                            for email in found_emails[:2]:
                                email_lower = email.lower()
                                if email_lower not in self.found_emails:
                                    new_emails.append(f"{email} (найден)")
                                    self.found_emails.add(email_lower)
                            
                            if new_emails:
                                item['email'] = '; '.join(new_emails)
                
                all_final_data = []
                for item in items_without_websites:
                    all_final_data.append(item)
                
                for group_info in domain_groups.values():
                    all_final_data.extend(group_info['items'])
            else:
                all_final_data = all_raw_data
            
            self.progress.emit("Обработка результатов...")
            self.progress_value.emit(90)
            
            final_results = []
            for item in all_final_data:

                if self.is_duplicate(item, final_results):
                    continue
                
                if item.get('email') or item.get('phone'):
                    final_results.append(item)
                    self.result_found.emit(item)
            
            self.progress_value.emit(100)
            self.progress.emit(f"Завершено. Найдено {len(final_results)} учреждений с контактами")
            self.search_completed.emit()
        
        except Exception as e:
            self.error_occurred.emit(str(e))
    
    def extract_basic_contact_info(self, item: Dict) -> Dict:
        contact_info = {
            'id': item.get('id', ''),
            'name': item.get('name', 'Не указано'),
            'email': '',
            'phone': '',
            'address': '',
            'website': '',
            'schedule': '',
            'lat': '',
            'lon': ''
        }
        
        emails = []
        phones = []
        websites = []
        
        if 'point' in item:
            contact_info['lat'] = str(item['point'].get('lat', ''))
            contact_info['lon'] = str(item['point'].get('lon', ''))
        
        if 'address' in item and item['address']:
            address = item['address']
            address_text = address.get('name', '')
            if not address_text:
                address_text = address.get('address_name', '')
            if not address_text and 'components' in address:
                components = []
                for comp in address['components']:
                    if 'street' in comp:
                        components.append(comp['street'])
                    if 'number' in comp:
                        components.append(comp['number'])
                if components:
                    address_text = ', '.join(components)
            contact_info['address'] = address_text
        
        if not contact_info['address'] and 'address_name' in item:
            contact_info['address'] = item['address_name']
        
        if 'contact_groups' in item:
            for group in item['contact_groups']:
                for contact in group.get('contacts', []):
                    contact_type = contact.get('type', '')
                    contact_value = contact.get('value', '')
                    if contact_type == 'email' and contact_value:
                        if contact_value not in emails:
                            emails.append(contact_value)
                    elif contact_type == 'phone' and contact_value:
                        phone_text = contact.get('print_text', contact_value)
                        comment = contact.get('comment', '')
                        if comment:
                            phone_text = f"{phone_text} ({comment})"
                        if phone_text not in phones:
                            phones.append(phone_text)
                    elif contact_type == 'website' and contact_value:
                        website_text = contact.get('print_text', contact.get('text', contact_value))
                        if website_text.startswith('http://link.2gis.ru/'):
                            website_text = contact.get('text', website_text)
                        if not any(social in website_text.lower() for social in [
                            'vk.com', 'ok.ru', 't.me', 'telegram', 'instagram', 
                            'facebook', 'youtube.com', 'twitter.com'
                        ]):
                            if website_text not in websites:
                                websites.append(website_text)
        
        contact_info['email'] = '; '.join(emails)
        contact_info['phone'] = '; '.join(phones[:3])
        contact_info['website'] = '; '.join(websites[:2])
        
        if 'schedule' in item and item['schedule']:
            schedule = item['schedule']
            if 'comment' in schedule:
                contact_info['schedule'] = schedule['comment']
            elif 'working_hours' in schedule:
                hours = []
                days_order = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
                for day in days_order:
                    if day in schedule and 'working_hours' in schedule[day]:
                        day_hours = schedule[day]['working_hours']
                        if day_hours:
                            wh = day_hours[0]
                            if 'from' in wh and 'to' in wh:
                                hours.append(f"{day}: {wh['from']}-{wh['to']}")
                if hours:
                    contact_info['schedule'] = '; '.join(hours[:2])
        
        return contact_info
    
    def is_duplicate(self, new_item, existing_items):
        new_email = new_item.get('email', '').strip()
        new_website = new_item.get('website', '').strip()
        
        for existing in existing_items:
            if new_email and existing.get('email'):
                new_emails = set([e.strip().lower() for e in new_email.split(';') if e.strip()])
                existing_emails = set([e.strip().lower() for e in existing['email'].split(';') if e.strip()])
                if new_emails & existing_emails:
                    return True
            
            if new_website and existing.get('website'):
                if self.normalize_website(new_website) == self.normalize_website(existing['website']):
                    return True
        
        return False
    
    def normalize_website(self, site):
        site = site.lower().strip()
        for prefix in ['https://', 'http://', 'www.']:
            if site.startswith(prefix):
                site = site[len(prefix):]
        return site.rstrip('/')
    
    def stop(self):
        self.is_running = False
    
    def get_city_coordinates(self, city_name: str) -> tuple:
        url = f"{self.base_url}/items/geocode"
        params = {
            'q': city_name,
            'fields': 'items.point',
            'key': self.api_key
        }
        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            data = response.json()
            if 'result' in data and 'items' in data['result'] and data['result']['items']:
                for item in data['result']['items']:
                    if 'point' in item:
                        return item['point']['lon'], item['point']['lat']
        except Exception as e:
            print(f"Ошибка геокодирования: {e}")
        return None, None
    
    def search_medical_organizations(self, lon: float, lat: float, radius_km: int, category: str) -> List[Dict]:
        radius_meters = radius_km * 1000
        url = f"{self.base_url}/items"
        params = {
            'q': category,
            'location': f"{lon},{lat}",
            'radius': radius_meters,
            'fields': 'items.point,items.schedule,items.contact_groups,items.address,items.external_content',
            'page_size': 50,
            'key': "rutnpt3272",
            'type': 'branch,building',
            'locale': 'ru_RU'
        }
        all_items = []
        page = 1
        while True:
            if not self.is_running:
                break
            params['page'] = page
            try:
                response = requests.get(url, params=params, timeout=10)
                response.raise_for_status()
                data = response.json()
                if 'result' in data and 'items' in data['result']:
                    items = data['result']['items']
                    if not items:
                        break
                    all_items.extend(items)
                    total = data['result'].get('total', 0)
                    if len(all_items) >= total:
                        break
                    page += 1
                    time.sleep(0.1)
                else:
                    break
            except Exception as e:
                print(f"Ошибка поиска на странице {page}: {e}")
                break
        return all_items    
class HospitalEmailFinderQt(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings_file = "api_settings.json"
        self.hospitals_data = []
        self.filtered_data = []
        self.search_worker = None
        self.medical_categories = [
            "Заполните список"
        ]
        self.init_ui()
        self.load_settings()
    def init_ui(self):
        self.setWindowTitle("2GIS Finder")
        self.setGeometry(100, 100, 1200, 800)
        self.set_minimalist_style()
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(10, 10, 10, 10)
        self.create_api_section(main_layout)
        self.create_search_section(main_layout)
        self.create_progress_section(main_layout)
        self.create_filter_section(main_layout)
        self.create_table_section(main_layout)
        self.create_status_bar(main_layout)
    def set_minimalist_style(self):
        style = """
        QMainWindow {
            background-color: #ffffff;
        }
        QLabel {
            color: #333333;
            font-size: 12px;
        }
        QLineEdit {
            padding: 6px;
            border: 1px solid #e0e0e0;
            border-radius: 2px;
            font-size: 12px;
            background-color: white;
        }
        QLineEdit:focus {
            border: 1px solid #999999;
        }
        QPushButton {
            padding: 6px 12px;
            border: 1px solid #e0e0e0;
            border-radius: 2px;
            font-size: 12px;
            background-color: #fafafa;
            color: #333333;
            min-width: 80px;
        }
        QPushButton:hover {
            background-color: #f5f5f5;
            border: 1px solid #cccccc;
        }
        QPushButton:pressed {
            background-color: #eeeeee;
        }
        QPushButton:disabled {
            background-color: #f5f5f5;
            color: #999999;
            border: 1px solid #e0e0e0;
        }
        QPushButton#primary {
            background-color: #666666;
            color: white;
            border: 1px solid #666666;
        }
        QPushButton#primary:hover {
            background-color: #555555;
            border: 1px solid #555555;
        }
        QPushButton#primary:pressed {
            background-color: #444444;
        }
        QGroupBox {
            font-size: 12px;
            font-weight: normal;
            border: 1px solid #e0e0e0;
            border-radius: 2px;
            margin-top: 8px;
            padding-top: 12px;
            background-color: #fafafa;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 8px;
            padding: 0 4px 0 4px;
            color: #666666;
        }
        QTableWidget {
            border: 1px solid #e0e0e0;
            border-radius: 2px;
            background-color: white;
            gridline-color: #f0f0f0;
            font-size: 12px;
        }
        QTableWidget::item {
            padding: 4px;
            border: none;
        }
        QTableWidget::item:selected {
            background-color: #f5f5f5;
            color: #333333;
        }
        QHeaderView::section {
            background-color: #fafafa;
            border: none;
            border-right: 1px solid #e0e0e0;
            border-bottom: 1px solid #e0e0e0;
            padding: 6px;
            font-weight: normal;
            color: #666666;
        }
        QComboBox {
            padding: 5px;
            border: 1px solid #e0e0e0;
            border-radius: 2px;
            font-size: 12px;
            background-color: white;
            min-width: 120px;
        }
        QComboBox:focus {
            border: 1px solid #999999;
        }
        QSpinBox {
            padding: 5px;
            border: 1px solid #e0e0e0;
            border-radius: 2px;
            font-size: 12px;
            background-color: white;
        }
        QProgressBar {
            border: 1px solid #e0e0e0;
            border-radius: 2px;
            text-align: center;
            height: 16px;
            background-color: #f5f5f5;
            font-size: 11px;
        }
        QProgressBar::chunk {
            background-color: #666666;
            border-radius: 1px;
        }
        QCheckBox {
            font-size: 12px;
            color: #333333;
        }
        QCheckBox::indicator {
            width: 14px;
            height: 14px;
        }
        QMenu {
            background-color: white;
            border: 1px solid #e0e0e0;
            padding: 4px;
        }
        QMenu::item {
            padding: 4px 20px;
            font-size: 12px;
        }
        QMenu::item:selected {
            background-color: #f5f5f5;
        }
        """
        self.setStyleSheet(style)
    def create_api_section(self, layout):
        api_group = QGroupBox("Настройки API")
        api_layout = QVBoxLayout()
        key_layout = QHBoxLayout()
        key_layout.addWidget(QLabel("API ключ:"))
        self.api_key_input = QLineEdit()
        self.api_key_input.setEchoMode(QLineEdit.Password)
        self.api_key_input.setPlaceholderText("Введите ваш API ключ 2ГИС")
        key_layout.addWidget(self.api_key_input)
        self.show_key_checkbox = QCheckBox("Показать")
        self.show_key_checkbox.toggled.connect(self.toggle_key_visibility)
        key_layout.addWidget(self.show_key_checkbox)
        save_key_btn = QPushButton("Сохранить")
        save_key_btn.clicked.connect(self.save_settings)
        key_layout.addWidget(save_key_btn)
        api_layout.addLayout(key_layout)
        info_label = QLabel("Получить API ключ: https://dev.2gis.ru/")
        info_label.setStyleSheet("color: #999999; font-size: 11px;")
        api_layout.addWidget(info_label)
        api_group.setLayout(api_layout)
        layout.addWidget(api_group)
    def create_search_section(self, layout):
        search_group = QGroupBox("Параметры поиска")
        search_layout = QVBoxLayout()
        first_row = QHBoxLayout()
        first_row.addWidget(QLabel("Город:"))
        self.city_input = QLineEdit()
        self.city_input.setPlaceholderText("Например: Москва")
        self.city_input.setText("Москва")
        self.city_input.setMaximumWidth(200)
        first_row.addWidget(self.city_input)
        first_row.addWidget(QLabel("Радиус (км):"))
        self.radius_spin = QSpinBox()
        self.radius_spin.setMinimum(1)
        self.radius_spin.setMaximum(50)
        self.radius_spin.setValue(10)
        self.radius_spin.setMaximumWidth(60)
        first_row.addWidget(self.radius_spin)
        edit_categories_btn = QPushButton("Категории поиска")
        edit_categories_btn.clicked.connect(self.edit_categories)
        first_row.addWidget(edit_categories_btn)
        
        
        first_row.addStretch()
        second_row = QHBoxLayout()
        self.search_btn = QPushButton("Начать поиск")
        self.search_btn.clicked.connect(self.start_search)
        self.search_btn.setObjectName("primary")
        second_row.addWidget(self.search_btn)
        self.stop_btn = QPushButton("Остановить")
        self.stop_btn.clicked.connect(self.stop_search)
        self.stop_btn.setEnabled(False)
        second_row.addWidget(self.stop_btn)
        clear_btn = QPushButton("Очистить")
        clear_btn.clicked.connect(self.clear_results)
        self.parse_websites_checkbox = QCheckBox("Искать email на сайтах")
        self.parse_websites_checkbox.setChecked(True)
        second_row.addWidget(clear_btn)
        second_row.addWidget(self.parse_websites_checkbox)        
        
        second_row.addStretch()
        search_layout.addLayout(first_row)
        search_layout.addLayout(second_row)
        search_group.setLayout(search_layout)
        layout.addWidget(search_group)
    def create_progress_section(self, layout):
        self.progress_label = QLabel("")
        self.progress_label.setStyleSheet("color: #666666; font-size: 11px;")
        layout.addWidget(self.progress_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
    def create_filter_section(self, layout):
        filter_group = QGroupBox("Фильтры и сортировка")
        filter_layout = QHBoxLayout()
        self.email_filter = QCheckBox("Только с email")
        self.email_filter.toggled.connect(self.apply_filters)
        filter_layout.addWidget(self.email_filter)
        self.phone_filter = QCheckBox("Только с телефоном")
        self.phone_filter.toggled.connect(self.apply_filters)
        filter_layout.addWidget(self.phone_filter)
        self.website_filter = QCheckBox("Только с сайтом")
        self.website_filter.toggled.connect(self.apply_filters)
        filter_layout.addWidget(self.website_filter)
        filter_layout.addStretch()
        filter_layout.addWidget(QLabel("Сортировка:"))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["По названию", "По адресу", "По email", "По категории"])
        self.sort_combo.currentIndexChanged.connect(self.apply_sorting)
        filter_layout.addWidget(self.sort_combo)
        filter_layout.addWidget(QLabel("Поиск:"))
        self.search_table_input = QLineEdit()
        self.search_table_input.setPlaceholderText("Поиск по таблице...")
        self.search_table_input.setMaximumWidth(200)
        self.search_table_input.textChanged.connect(self.search_in_table)
        filter_layout.addWidget(self.search_table_input)
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
    def create_table_section(self, layout):
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Название", "Email", "Телефон", "Адрес", "Сайт", "Режим работы", "Категория"
        ])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSortingEnabled(True)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        self.table.cellDoubleClicked.connect(self.on_cell_double_click)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        layout.addWidget(self.table)
    def on_cell_double_click(self, row, column):
        if column == 4:
            website_item = self.table.item(row, column)
            if website_item and website_item.text():
                self.open_website(website_item.text())
        elif column == 1:
            email_item = self.table.item(row, column)
            if email_item and email_item.text():
                import urllib.parse
                emails = [e.strip() for e in email_item.text().split(';')]
                if emails and emails[0]:
                    webbrowser.open(f"mailto:{emails[0]}")
    def create_status_bar(self, layout):
        status_layout = QHBoxLayout()
        self.status_label = QLabel("Готов к поиску")
        self.status_label.setStyleSheet("color: #666666; font-size: 11px;")
        status_layout.addWidget(self.status_label)
        status_layout.addStretch()
        export_excel_btn = QPushButton("Экспорт в Excel")
        export_excel_btn.clicked.connect(self.export_to_excel)
        status_layout.addWidget(export_excel_btn)
        export_csv_btn = QPushButton("Экспорт в CSV")
        export_csv_btn.clicked.connect(self.export_to_csv)
        status_layout.addWidget(export_csv_btn)
        layout.addLayout(status_layout)
    def show_context_menu(self, position):
        item = self.table.itemAt(position)
        if item is None:
            return
        row = item.row()
        menu = QMenu(self)
        if self.filtered_data:
            hospital_data = self.filtered_data[row] if row < len(self.filtered_data) else None
        else:
            hospital_data = self.hospitals_data[row] if row < len(self.hospitals_data) else None
        if not hospital_data:
            return
        open_2gis_action = QAction("Открыть в 2ГИС", self)
        open_2gis_action.triggered.connect(lambda: self.open_in_2gis(hospital_data))
        menu.addAction(open_2gis_action)
        website = hospital_data.get('website', '')
        if website:
            open_website_action = QAction("Открыть сайт", self)
            open_website_action.triggered.connect(lambda: self.open_website(website))
            menu.addAction(open_website_action)
        menu.addSeparator()
        email = hospital_data.get('email', '')
        if email:
            copy_email_action = QAction("Копировать email", self)
            copy_email_action.triggered.connect(lambda: QApplication.clipboard().setText(email))
            menu.addAction(copy_email_action)
        phone = hospital_data.get('phone', '')
        if phone:
            copy_phone_action = QAction("Копировать телефон", self)
            copy_phone_action.triggered.connect(lambda: QApplication.clipboard().setText(phone))
            menu.addAction(copy_phone_action)
        address = hospital_data.get('address', '')
        if address:
            copy_address_action = QAction("Копировать адрес", self)
            copy_address_action.triggered.connect(lambda: QApplication.clipboard().setText(address))
            menu.addAction(copy_address_action)
        if website:
            copy_website_action = QAction("Копировать сайт", self)
            copy_website_action.triggered.connect(lambda: QApplication.clipboard().setText(website))
            menu.addAction(copy_website_action)
        menu.exec_(self.table.mapToGlobal(position))
    def open_in_2gis(self, hospital_data):
        if hospital_data.get('lat') and hospital_data.get('lon'):
            url = f"https://2gis.ru/search/{hospital_data['name']}/geo/{hospital_data['lon']}%2C{hospital_data['lat']}"
            webbrowser.open(url)
        else:
            url = f"https://2gis.ru/search/{hospital_data['name']}"
            webbrowser.open(url)
    def open_website(self, website):
        websites = [w.strip() for w in website.split(';')]
        for site in websites:
            site = site.strip()
            if not site:
                continue
            if not site.startswith(('http://', 'https://')):
                if '://' in site:
                    webbrowser.open(site)
                else:
                    webbrowser.open(f"https://{site}")
            else:
                webbrowser.open(site)
            break
    def edit_categories(self):
        dialog = CategoryEditDialog(self.medical_categories, self)
        if dialog.exec_():
            self.medical_categories = dialog.get_categories()
            self.save_settings()
    def toggle_key_visibility(self, checked):
        if checked:
            self.api_key_input.setEchoMode(QLineEdit.Normal)
        else:
            self.api_key_input.setEchoMode(QLineEdit.Password)
    def save_settings(self):
        settings = {
            'api_key': self.api_key_input.text(),
            'categories': self.medical_categories
        }
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "Успех", "Настройки сохранены")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить настройки: {e}")
    def load_settings(self):
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.api_key_input.setText(settings.get('api_key', ''))
                    if 'categories' in settings:
                        self.medical_categories = settings['categories']
        except Exception as e:
            print(f"Ошибка загрузки настроек: {e}")
    def start_search(self):
        api_key = self.api_key_input.text().strip()
        city = self.city_input.text().strip()
        
        if not api_key:
            QMessageBox.warning(self, "Предупреждение", "Введите API ключ")
            return
        
        if not city:
            QMessageBox.warning(self, "Предупреждение", "Введите название города")
            return
        
        self.clear_results()
        self.search_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        self.search_worker = SearchWorker(api_key, city, self.radius_spin.value(), self.medical_categories)
    
        # Настройки email поиска
        self.search_worker.parse_websites = self.parse_websites_checkbox.isChecked()
    

        self.search_worker.smart_parsing = True
        
        self.search_worker.progress.connect(self.update_progress)
        self.search_worker.progress_value.connect(self.update_progress_value)
        self.search_worker.result_found.connect(self.add_result)
        self.search_worker.search_completed.connect(self.search_completed)
        self.search_worker.error_occurred.connect(self.handle_error)
        self.search_worker.start()        
    def stop_search(self):
        if self.search_worker:
            self.search_worker.stop()
            self.search_worker.wait()
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.progress_label.setText("Поиск остановлен")
    def clear_results(self):
        self.table.setRowCount(0)
        self.hospitals_data = []
        self.filtered_data = []
        self.status_label.setText("Готов к поиску")
        self.progress_label.setText("")
        self.progress_bar.setValue(0)
    def update_progress(self, message):
        self.progress_label.setText(message)
    def update_progress_value(self, value):
        self.progress_bar.setValue(value)
    def add_result(self, hospital_info):
        if hospital_info.get('id'):
            for existing in self.hospitals_data:
                if existing.get('id') == hospital_info['id']:
                    return
        
        hospital_website = hospital_info.get('website', '').strip()
        hospital_email = hospital_info.get('email', '').strip()
        hospital_phone = hospital_info.get('phone', '').strip()
        
        if hospital_email:
            new_emails = set([e.strip().lower() for e in hospital_email.split(';') if e.strip()])
            
            for existing in self.hospitals_data:
                existing_email = existing.get('email', '').strip()
                if existing_email:
                    existing_emails = set([e.strip().lower() for e in existing_email.split(';') if e.strip()])
                    if new_emails & existing_emails: 
                        return
        
        if hospital_website:
            def normalize_website(site):
                site = site.lower()
                for prefix in ['https://', 'http://', 'www.']:
                    if site.startswith(prefix):
                        site = site[len(prefix):]
                return site.rstrip('/')
            
            normalized_new = normalize_website(hospital_website)
            
            for existing in self.hospitals_data:
                existing_website = existing.get('website', '').strip()
                if existing_website:
                    normalized_existing = normalize_website(existing_website)
                    if normalized_new == normalized_existing:
                        return
        
        if not hospital_email and not hospital_website:
            def get_base_name(name):
                base = name.split(',')[0].strip()
                return base.lower()
            
            hospital_base_name = get_base_name(hospital_info['name'])
            
            for existing in self.hospitals_data:
                if not existing.get('email') and not existing.get('website'):
                    existing_base_name = get_base_name(existing['name'])
                    
                    if (hospital_base_name == existing_base_name and
                        existing.get('lat') and existing.get('lon') and 
                        hospital_info.get('lat') and hospital_info.get('lon')):
                        lat_diff = abs(float(existing['lat']) - float(hospital_info['lat']))
                        lon_diff = abs(float(existing['lon']) - float(hospital_info['lon']))
                        if lat_diff < 0.001 and lon_diff < 0.001:
                            return
        
        self.hospitals_data.append(hospital_info)
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        
        self.table.setItem(row_position, 0, QTableWidgetItem(hospital_info['name']))
        self.table.setItem(row_position, 1, QTableWidgetItem(hospital_info['email']))
        self.table.setItem(row_position, 2, QTableWidgetItem(hospital_info['phone']))
        self.table.setItem(row_position, 3, QTableWidgetItem(hospital_info['address']))
        self.table.setItem(row_position, 4, QTableWidgetItem(hospital_info['website']))
        self.table.setItem(row_position, 5, QTableWidgetItem(hospital_info['schedule']))
        self.table.setItem(row_position, 6, QTableWidgetItem(hospital_info.get('category', '')))
        
        self.status_label.setText(f"Найдено: {len(self.hospitals_data)} учреждений")    
    def search_completed(self):
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.progress_label.setText("Поиск завершен")
        self.progress_bar.setVisible(False)
        if self.hospitals_data:
            QMessageBox.information(
                self, 
                "Результат", 
                f"Найдено {len(self.hospitals_data)} учреждений с контактными данными"
            )
        else:
            QMessageBox.warning(
                self,
                "Результат",
                "Учреждения с контактными данными не найдены"
            )
    def handle_error(self, error_message):
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "Ошибка", error_message)
    def apply_filters(self):
        self.filtered_data = self.hospitals_data.copy()
        if self.email_filter.isChecked():
            self.filtered_data = [h for h in self.filtered_data if h['email']]
        if self.phone_filter.isChecked():
            self.filtered_data = [h for h in self.filtered_data if h['phone']]
        if self.website_filter.isChecked():
            self.filtered_data = [h for h in self.filtered_data if h['website']]
        self.update_table()
    def apply_sorting(self):
        if not self.filtered_data:
            self.filtered_data = self.hospitals_data.copy()
        sort_index = self.sort_combo.currentIndex()
        if sort_index == 0:
            self.filtered_data.sort(key=lambda x: x['name'])
        elif sort_index == 1:
            self.filtered_data.sort(key=lambda x: x['address'])
        elif sort_index == 2:
            self.filtered_data.sort(key=lambda x: x['email'])
        elif sort_index == 3:
            self.filtered_data.sort(key=lambda x: x.get('category', ''))
        self.update_table()
    def search_in_table(self, text):
        if not text:
            self.apply_filters()
            return
        text = text.lower()
        self.filtered_data = [
            h for h in self.hospitals_data
            if text in h['name'].lower() or
            text in h['email'].lower() or
            text in h['phone'].lower() or
            text in h['address'].lower() or
            text in h.get('category', '').lower()
        ]
        self.update_table()
    def update_table(self):
        self.table.setRowCount(0)
        data_to_show = self.filtered_data if self.filtered_data else self.hospitals_data
        for hospital in data_to_show:
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            self.table.setItem(row_position, 0, QTableWidgetItem(hospital['name']))
            self.table.setItem(row_position, 1, QTableWidgetItem(hospital['email']))
            self.table.setItem(row_position, 2, QTableWidgetItem(hospital['phone']))
            self.table.setItem(row_position, 3, QTableWidgetItem(hospital['address']))
            self.table.setItem(row_position, 4, QTableWidgetItem(hospital['website']))
            self.table.setItem(row_position, 5, QTableWidgetItem(hospital['schedule']))
            self.table.setItem(row_position, 6, QTableWidgetItem(hospital.get('category', '')))
        self.status_label.setText(f"Отображено: {len(data_to_show)} из {len(self.hospitals_data)}")
    def export_to_excel(self):
        if not self.hospitals_data:
            QMessageBox.warning(self, "Предупреждение", "Нет данных для экспорта")
            return
        filename, _ = QFileDialog.getSaveFileName(
            self, 
            "Сохранить как Excel",
            f"hospitals_{self.city_input.text()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel Files (*.xlsx)"
        )
        if filename:
            df = pd.DataFrame(self.hospitals_data)
            columns_map = {
                'name': 'Название',
                'email': 'Email',
                'phone': 'Телефон',
                'address': 'Адрес',
                'website': 'Сайт',
                'schedule': 'Режим работы',
                'category': 'Категория'
            }
            df = df[['name', 'email', 'phone', 'address', 'website', 'schedule', 'category']]
            df.columns = [columns_map[col] for col in df.columns]
            try:
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Учреждения', index=False)
                QMessageBox.information(self, "Успех", f"Данные сохранены в:\n{filename}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {e}")
    def export_to_csv(self):
        if not self.hospitals_data:
            QMessageBox.warning(self, "Предупреждение", "Нет данных для экспорта")
            return
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить как CSV",
            f"hospitals_{self.city_input.text()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV Files (*.csv)"
        )
        if filename:
            df = pd.DataFrame(self.hospitals_data)
            columns_map = {
                'name': 'Название',
                'email': 'Email',
                'phone': 'Телефон',
                'address': 'Адрес',
                'website': 'Сайт',
                'schedule': 'Режим работы',
                'category': 'Категория'
            }
            df = df[['name', 'email', 'phone', 'address', 'website', 'schedule', 'category']]
            df.columns = [columns_map[col] for col in df.columns]
            try:
                df.to_csv(filename, index=False, encoding='utf-8-sig')
                QMessageBox.information(self, "Успех", f"Данные сохранены в:\n{filename}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {e}")


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = HospitalEmailFinderQt()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()