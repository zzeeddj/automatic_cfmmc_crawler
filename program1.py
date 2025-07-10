import sys
import os
import datetime as dt
import json
from io import BytesIO
from typing import Sequence, List, Dict, Optional
from PIL import Image
from bs4 import BeautifulSoup
from requests import session
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget,
                             QTableWidgetItem, QPushButton, QLabel, QLineEdit, QDateEdit, QComboBox,
                             QCheckBox, QGroupBox, QDialog, QProgressBar, QMessageBox, QFileDialog,
                             QHeaderView, QAbstractItemView, QInputDialog, QAction)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate
from PyQt5.QtGui import QPixmap, QImage
import ddddocr  # 导入ddddocr库


class UserNamePasswordError(ValueError):
    pass


class VerificationCodeError(ValueError):
    pass


class CaptchaDialog(QDialog):
    """验证码输入对话框"""

    def __init__(self, captcha_image: bytes, parent=None):
        super().__init__(parent)
        self.setWindowTitle("验证码输入")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setFixedSize(300, 200)

        layout = QVBoxLayout()

        # 显示验证码图片
        self.captcha_label = QLabel()
        pixmap = QPixmap()
        pixmap.loadFromData(captcha_image)
        self.captcha_label.setPixmap(pixmap)
        layout.addWidget(self.captcha_label, alignment=Qt.AlignCenter)

        # 验证码输入框
        self.code_edit = QLineEdit()
        self.code_edit.setPlaceholderText("请输入验证码")
        layout.addWidget(self.code_edit)

        # 按钮区域
        btn_layout = QHBoxLayout()
        self.confirm_btn = QPushButton("确定")
        self.confirm_btn.clicked.connect(self.accept)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.confirm_btn)
        btn_layout.addWidget(self.cancel_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def get_code(self) -> str:
        return self.code_edit.text().strip()


class CFMMCCrawler(object):
    # 期货结算单下载器核心类
    base_url = "https://investorservice.cfmmc.com"
    login_url = base_url + '/login.do'
    logout_url = base_url + '/logout.do'
    data_url = base_url + '/customer/setParameter.do'
    excel_daily_download_url = base_url + '/customer/setupViewCustomerDetailFromCompanyWithExcel.do'
    excel_monthly_download_url = base_url + '/customer/setupViewCustomerMonthDetailFromCompanyWithExcel.do'
    header = {
        'Connection': 'keep-alive',
        'User-Agent': "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36",
    }
    query_type_dict = {'逐日': 'day', '逐笔': 'trade'}

    def __init__(self, division_name: str, company_short: str,
                 account_no: str, password: str,
                 output_dir: str, tushare_token: str) -> None:
        self.division_name, self.company_short = division_name, company_short
        self.account_no, self.password = account_no, password
        self.output_dir = output_dir
        self.tushare_token = tushare_token
        self._ss = None
        self.token = None
        self.ocr = ddddocr.DdddOcr()  # 初始化ddddocr

    def get_login_page(self) -> (str, bytes):
        """获取登录页面和验证码"""
        self._ss = session()
        res = self._ss.get(self.login_url, headers=self.header)
        bs = BeautifulSoup(res.text, features="lxml")
        token = bs.body.form.input['value']
        verification_code_url = self.base_url + bs.body.form.img['src']
        captcha_image = self._ss.get(verification_code_url).content
        return token, captcha_image

    def login(self, verification_code: str, token: str) -> None:
        """使用验证码登录"""
        post_data = {
            "org.apache.struts.taglib.html.TOKEN": token,
            "showSaveCookies": '',
            "userID": self.account_no,
            "password": self.password,
            "vericode": verification_code,
        }
        # 发送登录请求
        data_page = self._ss.post(self.login_url, data=post_data, headers=self.header, timeout=5)

        # 检查验证码错误
        if "验证码错误" in data_page.text:
            raise VerificationCodeError('登录失败, 验证码错误, 请重试!')
        # 检查用户名密码错误
        if '请勿在公用电脑上记录您的查询密码' in data_page.text:
            raise UserNamePasswordError(f"{self.company_short} 用户名密码错误!")

        # 获取登录token
        self.token = self._get_token(data_page.text)

    def logout(self) -> None:
        """登出"""
        if self.token:
            self._ss.post(self.logout_url)
            self.token = None

    def _check_args(self, query_type: str) -> None:
        if not self.token:
            raise RuntimeError('需要先登录成功才可进行查询!')
        if query_type not in self.query_type_dict.keys():
            raise ValueError('query_type 必须为 逐日 或 逐笔 !')

    def get_daily_data(self, date: dt.date, query_type: str) -> str:
        """下载日报数据"""
        self._check_args(query_type)

        trade_date = date.strftime('%Y-%m-%d')
        # 修改日报路径和文件名格式 - 按查询类型创建子目录
        path = os.path.join(self.output_dir, self.division_name, query_type)
        file_name = f"{self.division_name}-{self.company_short}_{date.strftime('%Y-%m-%d')}.xls"
        full_path = os.path.join(path, file_name)
        os.makedirs(path, exist_ok=True)

        post_data = {
            "org.apache.struts.taglib.html.TOKEN": self.token,
            "tradeDate": trade_date,
            "byType": self.query_type_dict[query_type]
        }
        data_page = self._ss.post(self.data_url, data=post_data, headers=self.header, timeout=5)
        self.token = self._get_token(data_page.text)

        self._download_file(self.excel_daily_download_url, full_path)
        return full_path

    def get_monthly_data(self, month: dt.date, query_type: str) -> str:
        """下载月报数据"""
        self._check_args(query_type)

        trade_date = month.strftime('%Y-%m')
        # 修改月报路径和文件名格式 - 按查询类型创建子目录
        path = os.path.join(self.output_dir, "月报", query_type)
        file_name = f"{self.division_name}-{self.company_short}_{month.strftime('%Y-%m')}.xls"
        full_path = os.path.join(path, file_name)
        os.makedirs(path, exist_ok=True)

        post_data = {
            "org.apache.struts.taglib.html.TOKEN": self.token,
            "tradeDate": trade_date,
            "byType": self.query_type_dict[query_type]
        }
        data_page = self._ss.post(self.data_url, data=post_data, headers=self.header, timeout=5)
        self.token = self._get_token(data_page.text)

        self._download_file(self.excel_monthly_download_url, full_path)
        return full_path

    @staticmethod
    def _get_token(page: str) -> str:
        token = BeautifulSoup(page, features="lxml").form.input['value']
        return token

    def _download_file(self, web_address: str, download_path: str) -> None:
        excel_response = self._ss.get(web_address)
        with open(download_path, 'wb') as fh:
            fh.write(excel_response.content)

    def get_trading_days(self, start_date: str, end_date: str) -> Sequence[dt.datetime]:
        """获取区间的交易日(周一至周五)"""
        start = dt.datetime.strptime(start_date, '%Y%m%d')
        end = dt.datetime.strptime(end_date, '%Y%m%d')

        trading_days = []
        current = start
        while current <= end:
            if current.weekday() < 5:  # 0-4 表示周一到周五
                trading_days.append(current)
            current += dt.timedelta(days=1)
        return trading_days

    @staticmethod
    def _generate_months_first_day(start_date: str, end_date: str) -> Sequence[dt.date]:
        """生成月份的第一天"""
        start = dt.date(int(start_date[:4]), int(start_date[4:6]), 1)
        end = dt.date(int(end_date[:4]), int(end_date[4:6]), 1)
        storage = []
        while start <= end:
            storage.append(start)
            next_month = start.month + 1
            next_year = start.year
            if next_month > 12:
                next_month = 1
                next_year += 1
            start = dt.date(next_year, next_month, 1)
        return storage


class DownloadThread(QThread):
    """下载线程"""
    progress_updated = pyqtSignal(int, str)
    finished = pyqtSignal()
    captcha_required = pyqtSignal(bytes)
    error_occurred = pyqtSignal(str)
    login_failed = pyqtSignal(str)  # 新增信号，用于登录失败通知

    def __init__(self, accounts: List[Dict], config: Dict, parent=None):
        super().__init__(parent)
        self.accounts = accounts
        self.config = config
        self.captcha_code = None
        self.cancelled = False
        self.max_retry = 1000  # 最大重试次数
        self.current_account_idx = 0
        self.current_account = None

    def run(self):
        total_accounts = len(self.accounts)

        for idx, account in enumerate(self.accounts):
            if self.cancelled:
                break

            self.current_account_idx = idx
            self.current_account = account

            # 创建下载器实例
            crawler = CFMMCCrawler(
                account['division_name'],
                account['company_short'],
                account['account_no'],
                account['password'],
                self.config['output_dir'],
                self.config.get('tushare_token', '')
            )

            try:
                # 登录流程
                login_success = False
                retry_count = 0

                while not login_success and retry_count < self.max_retry and not self.cancelled:
                    try:
                        # 获取验证码
                        token, captcha_image = crawler.get_login_page()

                        # 自动识别验证码
                        verification_code = crawler.ocr.classification(captcha_image)

                        # 尝试登录
                        crawler.login(verification_code, token)
                        login_success = True

                    except VerificationCodeError:
                        retry_count += 1
                        if retry_count >= self.max_retry:
                            self.login_failed.emit(f"{account['division_name']} 登录失败: 验证码错误次数过多")
                            self.captcha_required.emit(captcha_image)

                            # 等待验证码输入
                            while not self.captcha_code and not self.cancelled:
                                self.msleep(100)

                            if self.cancelled:
                                break

                            verification_code = self.captcha_code
                            self.captcha_code = None
                            break
                        continue
                    except UserNamePasswordError as e:
                        self.login_failed.emit(str(e))
                        break
                    except Exception as e:
                        self.login_failed.emit(f"{account['division_name']} 登录失败: {str(e)}")
                        break

                if not login_success or self.cancelled:
                    continue

                # 下载数据
                report_types = self.config['report_types']
                query_types = self.config['query_types']
                start_date = self.config['start_date']
                end_date = self.config['end_date']

                # 计算总任务数
                total_tasks = 0
                if '日报' in report_types:
                    total_tasks += len(query_types)
                if '月报' in report_types:
                    total_tasks += len(query_types)

                current_task = 0

                # 日报下载
                if '日报' in report_types:
                    trading_days = crawler.get_trading_days(start_date, end_date)
                    total_days = len(trading_days)

                    for query_type in query_types:
                        if self.cancelled:
                            break

                        for day_idx, date in enumerate(trading_days):
                            if self.cancelled:
                                break
                            try:
                                file_path = crawler.get_daily_data(date, query_type)
                                current_task_progress = (current_task * 100 + (
                                        day_idx + 1) / total_days * 100) / total_tasks
                                msg = f"{account['division_name']} - {date.strftime('%Y-%m-%d')} {query_type}报表下载完成"
                                self.progress_updated.emit(int(current_task_progress), msg)
                            except Exception as e:
                                self.error_occurred.emit(
                                    f"{account['division_name']} {date.strftime('%Y-%m-%d')} {query_type}报表下载失败: {str(e)}")

                        current_task += 1

                # 月报下载
                if '月报' in report_types:
                    months = crawler._generate_months_first_day(start_date, end_date)
                    total_months = len(months)

                    for query_type in query_types:
                        if self.cancelled:
                            break

                        for month_idx, month in enumerate(months):
                            if self.cancelled:
                                break
                            try:
                                file_path = crawler.get_monthly_data(month, query_type)
                                current_task_progress = (current_task * 100 + (
                                        month_idx + 1) / total_months * 100) / total_tasks
                                msg = f"{account['division_name']} - {month.strftime('%Y-%m')} {query_type}报表下载完成"
                                self.progress_updated.emit(int(current_task_progress), msg)
                            except Exception as e:
                                self.error_occurred.emit(
                                    f"{account['division_name']} {month.strftime('%Y-%m')} {query_type}报表下载失败: {str(e)}")

                        current_task += 1

                # 登出
                crawler.logout()

                # 更新进度
                account_progress = int((idx + 1) / total_accounts * 100)
                self.progress_updated.emit(account_progress, f"{account['division_name']} 下载完成")

            except Exception as e:
                self.error_occurred.emit(f"{account['division_name']} 下载失败: {str(e)}")
                continue  # 继续下一个账户

        self.finished.emit()

    def cancel(self):
        self.cancelled = True
        self.captcha_code = None

    def set_captcha(self, code: str):
        self.captcha_code = code


class AccountManager(QWidget):
    """账户管理界面"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.config_file = "config.json"
        self.config = self.load_config()
        self.original_accounts = self.config.get('accounts', [])  # 保存原始账户列表用于搜索
        self.current_accounts = self.original_accounts.copy()  # 当前显示的账户列表

        # 添加删除操作的撤销栈
        self.deleted_stack = []  # 存储被删除的账户列表

        # 创建UI
        self.init_ui()

    def load_config(self) -> Dict:
        """加载配置文件"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {'accounts': [], 'output_dir': './downloads'}
        return {'accounts': [], 'output_dir': './downloads'}

    def save_config(self):
        """保存配置文件"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def init_ui(self):
        main_layout = QVBoxLayout()

        # 搜索区域
        search_layout = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("输入事业部、公司简称或账户搜索...")
        self.search_edit.textChanged.connect(self.filter_accounts)
        search_layout.addWidget(self.search_edit)

        self.search_btn = QPushButton("搜索")
        self.search_btn.clicked.connect(self.filter_accounts)
        search_layout.addWidget(self.search_btn)

        self.clear_search_btn = QPushButton("清除")
        self.clear_search_btn.clicked.connect(self.clear_search)
        search_layout.addWidget(self.clear_search_btn)

        main_layout.addLayout(search_layout)

        # 账户表格
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(['选择', '事业部', '公司简称', '账号', '密码'])
        self.table.setColumnWidth(0, 60)
        self.table.setColumnWidth(1, 150)
        self.table.setColumnWidth(2, 150)
        self.table.setColumnWidth(3, 120)
        self.table.setColumnWidth(4, 150)

        # 设置表格属性
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.load_accounts_to_table()

        # 按钮区域
        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("添加账户")
        self.add_btn.clicked.connect(self.add_account)
        self.edit_btn = QPushButton("编辑账户")
        self.edit_btn.clicked.connect(self.edit_account)
        self.delete_btn = QPushButton("删除账户")
        self.delete_btn.clicked.connect(self.delete_account)
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.select_all_accounts)
        self.deselect_all_btn = QPushButton("取消全选")
        self.deselect_all_btn.clicked.connect(self.deselect_all_accounts)

        # 添加撤销按钮
        self.undo_delete_btn = QPushButton("撤销删除")
        self.undo_delete_btn.clicked.connect(self.undo_delete)
        self.undo_delete_btn.setEnabled(False)  # 初始不可用

        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(self.edit_btn)
        btn_layout.addWidget(self.delete_btn)
        btn_layout.addWidget(self.select_all_btn)
        btn_layout.addWidget(self.deselect_all_btn)
        btn_layout.addWidget(self.undo_delete_btn)

        # 下载设置区域
        settings_group = QGroupBox("下载设置")
        settings_layout = QVBoxLayout()

        # 日期范围
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("开始日期:"))
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setDate(dt.date.today() - dt.timedelta(days=30))
        self.start_date_edit.setDisplayFormat("yyyyMMdd")
        date_layout.addWidget(self.start_date_edit)

        date_layout.addWidget(QLabel("结束日期:"))
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDate(dt.date.today())
        self.end_date_edit.setDisplayFormat("yyyyMMdd")
        date_layout.addWidget(self.end_date_edit)
        settings_layout.addLayout(date_layout)

        # 报表类型选择 (多选)
        report_group = QGroupBox("报表类型")
        report_layout = QVBoxLayout()
        self.daily_check = QCheckBox("日报")
        self.daily_check.setChecked(True)
        self.monthly_check = QCheckBox("月报")
        self.monthly_check.setChecked(True)
        report_layout.addWidget(self.daily_check)
        report_layout.addWidget(self.monthly_check)
        report_group.setLayout(report_layout)

        # 查询类型选择 (多选)
        query_group = QGroupBox("查询类型")
        query_layout = QVBoxLayout()
        self.day_check = QCheckBox("逐日")
        self.day_check.setChecked(True)
        self.trade_check = QCheckBox("逐笔")
        self.trade_check.setChecked(True)
        query_layout.addWidget(self.day_check)
        query_layout.addWidget(self.trade_check)
        query_group.setLayout(query_layout)

        # 将报表和查询类型选择放入同一行
        type_layout = QHBoxLayout()
        type_layout.addWidget(report_group)
        type_layout.addWidget(query_group)
        settings_layout.addLayout(type_layout)

        # 输出目录
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(QLabel("输出目录:"))
        self.dir_edit = QLineEdit(self.config.get('output_dir', './downloads'))
        dir_layout.addWidget(self.dir_edit)
        self.browse_btn = QPushButton("浏览...")
        self.browse_btn.clicked.connect(self.browse_directory)
        dir_layout.addWidget(self.browse_btn)
        settings_layout.addLayout(dir_layout)

        settings_group.setLayout(settings_layout)

        # 下载按钮和进度条
        self.download_btn = QPushButton("开始下载")
        self.download_btn.clicked.connect(self.start_download)
        self.cancel_btn = QPushButton("取消下载")
        self.cancel_btn.clicked.connect(self.cancel_download)
        self.cancel_btn.setEnabled(False)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_label = QLabel("准备下载...")

        # 添加到主布局
        main_layout.addWidget(self.table)
        main_layout.addLayout(btn_layout)
        main_layout.addWidget(settings_group)
        main_layout.addWidget(self.download_btn)
        main_layout.addWidget(self.cancel_btn)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.progress_label)

        self.setLayout(main_layout)
        self.setWindowTitle("期货结算单下载器")
        self.resize(900, 700)

    def load_accounts_to_table(self):
        """加载账户到表格"""
        self.table.setRowCount(len(self.current_accounts))

        for i, account in enumerate(self.current_accounts):
            # 选择框
            chk = QCheckBox()
            chk.setChecked(True)
            self.table.setCellWidget(i, 0, chk)

            # 事业部
            self.table.setItem(i, 1, QTableWidgetItem(account['division_name']))

            # 公司简称
            self.table.setItem(i, 2, QTableWidgetItem(account['company_short']))

            # 账号
            self.table.setItem(i, 3, QTableWidgetItem(account['account_no']))

            # 密码 (明文显示)
            self.table.setItem(i, 4, QTableWidgetItem(account['password']))

    def filter_accounts(self):
        """根据搜索条件过滤账户"""
        search_text = self.search_edit.text().strip().lower()
        if not search_text:
            self.current_accounts = self.original_accounts.copy()
        else:
            self.current_accounts = [
                acc for acc in self.original_accounts
                if (search_text in acc['division_name'].lower() or
                    search_text in acc['company_short'].lower() or
                    search_text in acc['account_no'].lower())
            ]
        self.load_accounts_to_table()

    def clear_search(self):
        """清除搜索条件"""
        self.search_edit.clear()
        self.current_accounts = self.original_accounts.copy()
        self.load_accounts_to_table()

    def get_selected_accounts(self) -> List[Dict]:
        """获取选中的账户"""
        accounts = []
        for i in range(self.table.rowCount()):
            chk = self.table.cellWidget(i, 0)
            if chk and chk.isChecked():
                account = {
                    'division_name': self.table.item(i, 1).text(),
                    'company_short': self.table.item(i, 2).text(),
                    'account_no': self.table.item(i, 3).text(),
                    'password': self.table.item(i, 4).text()
                }
                accounts.append(account)
        return accounts

    def add_account(self):
        """添加新账户"""
        # 使用对话框获取账户信息
        division_name, ok = QInputDialog.getText(self, "添加账户", "事业部:")
        if not ok or not division_name:
            return

        company_short, ok = QInputDialog.getText(self, "添加账户", "公司简称:")
        if not ok or not company_short:
            return

        account_no, ok = QInputDialog.getText(self, "添加账户", "账号:")
        if not ok or not account_no:
            return

        password, ok = QInputDialog.getText(self, "添加账户", "密码:", QLineEdit.Normal, "")
        if not ok:
            return

        # 添加到原始账户列表
        new_account = {
            'division_name': division_name,
            'company_short': company_short,
            'account_no': account_no,
            'password': password
        }
        self.original_accounts.append(new_account)
        self.current_accounts.append(new_account)

        # 保存配置
        self.config['accounts'] = self.original_accounts
        self.save_config()

        # 刷新表格
        self.load_accounts_to_table()

    def edit_account(self):
        """编辑账户"""
        selected_row = self.table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "警告", "请先选择要编辑的账户!")
            return

        # 获取当前账户信息
        account = {
            'division_name': self.table.item(selected_row, 1).text(),
            'company_short': self.table.item(selected_row, 2).text(),
            'account_no': self.table.item(selected_row, 3).text(),
            'password': self.table.item(selected_row, 4).text()
        }

        # 使用对话框编辑账户信息
        division_name, ok = QInputDialog.getText(self, "编辑账户", "事业部:", QLineEdit.Normal,
                                                 account['division_name'])
        if not ok:
            return

        company_short, ok = QInputDialog.getText(self, "编辑账户", "公司简称:", QLineEdit.Normal,
                                                 account['company_short'])
        if not ok:
            return

        account_no, ok = QInputDialog.getText(self, "编辑账户", "账号:", QLineEdit.Normal, account['account_no'])
        if not ok:
            return

        password, ok = QInputDialog.getText(self, "编辑账户", "密码:", QLineEdit.Normal, account['password'])
        if not ok:
            return

        # 更新账户信息
        account['division_name'] = division_name
        account['company_short'] = company_short
        account['account_no'] = account_no
        account['password'] = password

        # 更新原始账户列表
        for i, acc in enumerate(self.original_accounts):
            if acc['account_no'] == account['account_no']:
                self.original_accounts[i] = account
                break

        # 更新当前显示的账户列表
        for i, acc in enumerate(self.current_accounts):
            if acc['account_no'] == account['account_no']:
                self.current_accounts[i] = account
                break

        # 保存配置
        self.config['accounts'] = self.original_accounts
        self.save_config()

        # 刷新表格
        self.load_accounts_to_table()

    def delete_account(self):
        """删除勾选的账户"""
        selected_accounts = self.get_selected_accounts()
        if not selected_accounts:
            QMessageBox.warning(self, "警告", "请先勾选要删除的账户!")
            return

        # 将被删除的账户保存到撤销栈
        self.deleted_stack.append(selected_accounts.copy())
        self.undo_delete_btn.setEnabled(True)  # 启用撤销按钮

        # 简化删除确认对话框
        reply = QMessageBox.question(self, "确认删除",
                                     f"确定要删除选中的{len(selected_accounts)}个账户吗?",
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)

        if reply == QMessageBox.Yes:
            try:
                # 获取要删除的账号列表
                account_nos = [acc['account_no'] for acc in selected_accounts]

                # 从原始账户列表中删除
                self.original_accounts = [acc for acc in self.original_accounts
                                          if acc['account_no'] not in account_nos]

                # 从当前显示的账户列表中删除
                self.current_accounts = [acc for acc in self.current_accounts
                                         if acc['account_no'] not in account_nos]

                # 保存配置
                self.config['accounts'] = self.original_accounts
                self.save_config()

                # 刷新表格
                self.load_accounts_to_table()

                # 显示删除成功提示
                QMessageBox.information(self, "提示", f"成功删除{len(selected_accounts)}个账户!")

            except Exception as e:
                QMessageBox.critical(self, "错误", f"删除账户时出错: {str(e)}")
        else:
            # 用户取消删除，从撤销栈中移除
            if self.deleted_stack:
                self.deleted_stack.pop()
                if not self.deleted_stack:
                    self.undo_delete_btn.setEnabled(False)

    def undo_delete(self):
        """撤销最近一次删除操作"""
        if not self.deleted_stack:
            QMessageBox.information(self, "提示", "没有可撤销的操作")
            return

        # 获取最近删除的账户列表
        last_deleted = self.deleted_stack.pop()

        # 检查撤销栈是否为空
        if not self.deleted_stack:
            self.undo_delete_btn.setEnabled(False)

        # 恢复被删除的账户
        for account in last_deleted:
            # 检查账户是否已存在
            exists = any(acc['account_no'] == account['account_no'] for acc in self.original_accounts)
            if not exists:
                self.original_accounts.append(account)
                self.current_accounts.append(account)

        # 保存配置
        self.config['accounts'] = self.original_accounts
        self.save_config()

        # 刷新表格
        self.load_accounts_to_table()

        # 显示撤销成功提示
        QMessageBox.information(self, "提示", f"已恢复{len(last_deleted)}个账户!")

    def select_all_accounts(self):
        """全选账户"""
        for i in range(self.table.rowCount()):
            chk = self.table.cellWidget(i, 0)
            if chk:
                chk.setChecked(True)

    def deselect_all_accounts(self):
        """取消全选账户"""
        for i in range(self.table.rowCount()):
            chk = self.table.cellWidget(i, 0)
            if chk:
                chk.setChecked(False)

    def browse_directory(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录", self.dir_edit.text())
        if dir_path:
            self.dir_edit.setText(dir_path)

    def start_download(self):
        """开始下载"""
        # 保存当前配置
        self.config['output_dir'] = self.dir_edit.text()
        self.save_config()

        # 获取选中的账户
        accounts = self.get_selected_accounts()
        if not accounts:
            QMessageBox.warning(self, "警告", "请至少选择一个账户！")
            return

        # 获取选中的报表类型
        report_types = []
        if self.daily_check.isChecked():
            report_types.append('日报')
        if self.monthly_check.isChecked():
            report_types.append('月报')

        if not report_types:
            QMessageBox.warning(self, "警告", "请至少选择一种报表类型！")
            return

        # 获取选中的查询类型
        query_types = []
        if self.day_check.isChecked():
            query_types.append('逐日')
        if self.trade_check.isChecked():
            query_types.append('逐笔')

        if not query_types:
            QMessageBox.warning(self, "警告", "请至少选择一种查询类型！")
            return

        # 准备下载配置
        download_config = {
            'start_date': self.start_date_edit.date().toString("yyyyMMdd"),
            'end_date': self.end_date_edit.date().toString("yyyyMMdd"),
            'report_types': report_types,
            'query_types': query_types,
            'output_dir': self.dir_edit.text(),
            'tushare_token': self.config.get('tushare_token', '')
        }

        # 创建下载线程
        self.download_thread = DownloadThread(accounts, download_config)
        self.download_thread.progress_updated.connect(self.update_progress)
        self.download_thread.finished.connect(self.download_finished)
        self.download_thread.captcha_required.connect(self.show_captcha_dialog)
        self.download_thread.login_failed.connect(self.show_error)
        self.download_thread.error_occurred.connect(self.show_error)

        # 更新UI状态
        self.download_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.progress_label.setText("开始下载...")

        # 启动下载线程
        self.download_thread.start()

    def cancel_download(self):
        """取消下载"""
        if hasattr(self, 'download_thread') and self.download_thread.isRunning():
            self.download_thread.cancel()
            self.cancel_btn.setEnabled(False)
            self.progress_label.setText("下载已取消")

    def download_finished(self):
        """下载完成"""
        self.download_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.progress_label.setText("下载完成")

    def update_progress(self, value: int, message: str):
        """更新进度"""
        self.progress_bar.setValue(value)
        self.progress_label.setText(message)

    def show_captcha_dialog(self, captcha_image: bytes):
        """显示验证码对话框"""
        self.captcha_dialog = CaptchaDialog(captcha_image, self)
        if self.captcha_dialog.exec_() == QDialog.Accepted:
            self.download_thread.set_captcha(self.captcha_dialog.get_code())
        else:
            self.download_thread.cancel()

    def show_error(self, message: str):
        """显示错误信息"""
        QMessageBox.critical(self, "错误", message)


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("期货结算单下载器")
        self.setGeometry(100, 100, 900, 700)

        self.account_manager = AccountManager()
        self.setCentralWidget(self.account_manager)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())