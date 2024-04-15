# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'gui.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QSlider, QLabel, QVBoxLayout, QHBoxLayout, QFrame, \
    QPushButton, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QImage, QPixmap, QColor, QTextCharFormat
from PyQt5.QtWidgets import QProgressDialog
from invoice import Ui_mainWindow
from PyQt5.QtCore import QCoreApplication
import time
from functools import partial
import pandas as pd
from xlsx_utils import *
from pdf_utils import *


class XlsxType(Enum):
    CLIENT = "client"
    DEVICE = "device"
    INVOICE = "invoice"


class DirectoryPathType(Enum):
    PDF = "pdf"
    SAVE = "save"


class CustomForm(QMainWindow, Ui_mainWindow):
    client_info_xlsx = None
    device_info_xlsx = None
    invoice_info_xlsx = None
    client_info_df: pd.DataFrame = None
    device_info_df: pd.DataFrame = None
    invoice_info_df: pd.DataFrame = None
    billing_info_df: pd.DataFrame = None
    year: int = None
    month: int = None
    pdf_dir_path: str = None
    save_dir_path: str = None

    def __init__(self):
        super().__init__()
        self.setupUi(self)  # 添加此行
        # 绑定函数
        # Connect the 'clicked' signal to the handler function
        try:
            self.pushButton.clicked.disconnect()
            self.pushButton_2.clicked.disconnect()
            self.pushButton_3.clicked.disconnect()
            self.pushButton_4.clicked.disconnect()
            self.pushButton_5.clicked.disconnect()
            self.pushButton_6.clicked.disconnect()



        except TypeError:
            pass  # Ignore if no connections found
        self.pushButton.clicked.connect(partial(self.on_xlsx_open_button_clicked, XlsxType.CLIENT))
        self.pushButton_2.clicked.connect(partial(self.on_xlsx_open_button_clicked, XlsxType.DEVICE))
        self.pushButton_3.clicked.connect(partial(self.on_xlsx_open_button_clicked, XlsxType.INVOICE))

        self.pushButton_6.clicked.connect(partial(self.on_dir_open_button_clicked, DirectoryPathType.PDF))
        self.pushButton_5.clicked.connect(partial(self.on_dir_open_button_clicked, DirectoryPathType.SAVE))

        self.pushButton_4.clicked.connect(self.process)

    def show_progress_dialog(self):
        progress_dialog = QProgressDialog("加载模型中...", "取消", 0, 100, self)
        progress_dialog.setWindowTitle("进度")
        progress_dialog.setWindowModality(QtCore.Qt.WindowModal)
        progress_dialog.show()

        for i in range(101):
            progress_dialog.setValue(i)
            QtCore.QThread.msleep(50)  # 模拟任务耗时

            if progress_dialog.wasCanceled():
                break

        progress_dialog.close()

    def on_dir_open_button_clicked(self, dir_type: DirectoryPathType):
        if dir_type == DirectoryPathType.PDF:

            self.pdf_dir_path = self.open_directory()
            if not self.pdf_dir_path:
                return
        elif dir_type == DirectoryPathType.SAVE:
            self.save_dir_path = self.open_directory()
            if not self.save_dir_path:
                return

    def on_xlsx_open_button_clicked(self, xlsx_type: XlsxType):
        if xlsx_type == XlsxType.CLIENT:
            self.client_info_xlsx = self.open_xlsx_file()
            if not self.client_info_xlsx:
                return
            self.client_info_df = pd.read_excel(self.client_info_xlsx)

        elif xlsx_type == XlsxType.DEVICE:
            self.device_info_xlsx = self.open_xlsx_file()
            if not self.device_info_xlsx:
                return
            self.billing_info_df = pd.read_excel(self.device_info_xlsx)
        elif xlsx_type == XlsxType.INVOICE:
            self.invoice_info_xlsx = self.open_xlsx_file()
            if not self.invoice_info_xlsx:
                return
            # 判断一下这个sheet有几个, 如果sheet不足两个，提示
            sheet_number = pd.ExcelFile(self.invoice_info_xlsx).sheet_names
            if len(sheet_number) < 2:
                QMessageBox.warning(self, "错误", "发票信息表格sheet不足两个")
                return

            self.device_info_df = pd.read_excel(self.invoice_info_xlsx, sheet_name=0)
            self.invoice_info_df = pd.read_excel(self.invoice_info_xlsx, sheet_name=1)

    def open_directory(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        folder_path = QFileDialog.getExistingDirectory(self, "打开文件夹", "", options=options)
        if folder_path:
            QMessageBox.information(self, "成功", f"成功打开文件夹：{folder_path}")
            return folder_path
        else:
            QMessageBox.warning(self, "错误", "未选择文件夹或文件夹打开失败")
            return None

    def insert_text(self, text: str, color: str):
        if color == "green":
            color = QColor(0, 255, 0)
        elif color == "red":
            color = QColor(255, 0, 0)
        else:
            color = QColor(0, 0, 0)
        text += "\n"
        cursor = self.statusTextEdit.textCursor()
        format = QTextCharFormat()
        format.setForeground(color)
        cursor.insertText(text, format)
        cursor.movePosition(cursor.End)
        self.statusTextEdit.setTextCursor(cursor)
        self.statusTextEdit.ensureCursorVisible()


    def process(self):

        # # 将光标移动到文档的末尾
        # # 在添加内容前
        # cursor = self.statusTextEdit.textCursor()
        #
        # # 创建一个QTextCharFormat对象并设置其颜色为绿色
        # format = QTextCharFormat()
        # format.setForeground(QColor(0, 255, 0))  # RGB for green
        #
        # # 使用QTextCursor的insertText方法插入文本，并使用QTextCharFormat设置文本的格式
        # cursor.insertText(f"开始选课: {course_name}\n", format)
        #
        # # 将光标移动到文档的末尾
        # cursor.movePosition(cursor.End)
        # self.statusTextEdit.setTextCursor(cursor)
        #
        # # 确保光标可见，这将导致textEdit滚动到底部
        # self.statusTextEdit.ensureCursorVisible()

        try:
            self.year = int(self.comboBox.currentText())
            self.month = int(self.comboBox_2.currentText())
            if self.year is None or self.month is None:
                QMessageBox.warning(self, "错误", "请先选择年份和月份")
                return
            if self.pdf_dir_path is None:
                QMessageBox.warning(self, "错误", "请先选择PDF文件夹")
                return
            if self.save_dir_path is None:
                QMessageBox.warning(self, "错误", "请先选择保存文件夹")
                return
            self.insert_text("开始处理处理发票, 数据加载完毕", "green")
            QApplication.processEvents()  # 处理事件队列

            filtered_df = filter_invoice_by_year_month(self.invoice_info_df, self.year, self.month)
            invoice_groups: dict[str, pd.DataFrame] = {}
            # group a dataframe by the column name "税控发票号1"
            for name, group in filtered_df.groupby("税控发票号1"):
                invoice_groups[name] = group
            self.insert_text(f"发票分组完成，共有{len(invoice_groups)}组", "green")
            QApplication.processEvents()  # 处理事件队列

            pdf_data = get_all_invoice_file(self.pdf_dir_path,
                                            FileType.PDF)
            service_template_df = pd.DataFrame(columns=['资产管理编号', '委托人', '委托单位', '委托单位地址', '委托人电话', '委托单位所属区域', '委托单位行业领域',
           '委托单位邮编', '委托单位邮箱', '委托单位传真', '委托单位开户行', '委托单位银行账号', '发票代码', '发票号码',
           '发票金额', '发票日期', '操作人姓名', '服务次数', '完成样品数', '使用机时', '收费（元）', '项目名称（代号）',
           '项目编号', '项目来源', '委托内容及检测要求', '委托人评价', '服务方式', '项目学科领域', '是否签订协议',
           '对外服务地址', '非适用简易程序海关《通知书》编号'])
            empty_info = {}
            for invoice_id, df in invoice_groups.items():
                has_exception = False
                device_id = None
                custom_info = None
                billing = None
                billing_info = None
                invoice_path = get_invoice_path(invoice_id, pdf_data)
                try:
                    device_id = get_deviceid_by_invoice_id(invoice_id, df, self.device_info_df)
                    billing_info = get_operator_and_billing(invoice_id, device_id, self.billing_info_df)
                    billing = extract_final_price(invoice_path)
                except Exception as e:
                    has_exception = True
                    print(e)
                try:
                    custom_info = get_custom_info_by_invoice_id(invoice_id, df, self.client_info_df)
                except Exception as e:
                    has_exception = True
                    print(e)
                if has_exception:
                    company_name = get_company_name_by_invoice_id(invoice_id, df, self.client_info_df)
                    if company_name not in empty_info:
                        empty_info[company_name] = [invoice_id]
                    empty_info[company_name].append(invoice_id)
                else:
                    pass
                    # service_template_df = pd.concat()
                date = str(df["凭证日期（开票）"].iloc[-1]).split(" ")[0]
                if billing == -1:
                    continue
                if custom_info is not None:
                    # 委托人 用英语
                    client_name = custom_info["委托人"]
                    client_unit = custom_info["委托单位"]
                    client_address = custom_info["委托单位地址"]
                    client_phone = custom_info["委托人电话"]
                    if (not pd.isna(list(custom_info.values())[3])) and (not pd.isna(list(custom_info.values())[4])):
                        client_region = "-".join(list(custom_info.values())[3:5])
                    else:
                        client_region = ""
                else:
                    client_name = ""
                    client_unit = ""
                    client_address = ""
                    client_phone = ""
                    client_region = ""
                if billing_info is not None:
                    operator_name = billing_info["操作员1姓名"]
                    serving_times = len(df)
                    # 获取收费标准 仪器参考收费标准
                    billing_criteria = str(billing_info["仪器参考收费标准"])
                    # 使用正则表达式匹配匹配里面的数字
                    # 使用正则表达式匹配匹配里面的数字
                    price_pattern = r'(\d+).*'
                    matches = re.findall(price_pattern, billing_criteria)
                    if matches:
                        # 将匹配到的最后一个价格转换为浮点数并返回
                        price = float(matches[0])
                    else:
                        price = -1
                    billing_criteria_str = f"{price}元/时"
                    using_hours = billing / price
                else:
                    operator_name = ""
                    serving_times = -1
                    billing_criteria_str = ""
                    using_hours = -1

                new_data = {"资产管理编号": device_id, "委托人": client_name, "委托单位": client_unit,
                            "委托单位地址": client_address, "委托人电话": client_phone, "委托单位所属区域": client_region,
                            "发票号码": invoice_id,
                            "发票金额": billing, "发票日期": date, "操作人姓名": operator_name, "服务次数": serving_times,
                            "收费标准": billing_criteria_str, "使用机时": using_hours,
                            "完成样品数": int(serving_times),
                            }
                service_template_df = pd.concat([service_template_df, pd.DataFrame([new_data])])
                self.insert_text(f"处理发票{invoice_id}完成", "green")
                QApplication.processEvents()  # 处理事件队列

                print("=====================================")

            service_template_df["项目来源"] = "其他"  # 给一个column全部赋值成这个可以这样写吗
            service_template_df["委托人评价"] = "好"
            service_template_df["服务方式"] = "技术共享"
            service_template_df["项目学科领域"] = "电子与通信技术"
            service_template_df["是否签订协议"] = "否"
            new_empty_df = pd.DataFrame(columns=["Company Name", "Invoice ID"])
            for k, v in empty_info.items():
                # 添加dataframe, 但是dataframe没有append方法，所以要用pd.DataFrame
                new_empty_df = pd.concat([new_empty_df, pd.DataFrame([[k, v]], columns=["Company Name", "Invoice ID"])])
            new_empty_df.to_excel("未找到信息的发票.xlsx", index=False)
            self.insert_text(f"处理完成，保存excel文件", "green")
            QApplication.processEvents()  # 处理事件队列

            self.save_excel(new_empty_df, service_template_df)
            self.insert_text(f"客户信息缺失值保存完毕", "green")
            self.insert_text(f"开始拆分发票", "green")
            QApplication.processEvents()  # 处理事件队列

            save_invoice_by_number(service_template_df, 50, self.save_dir_path, pdf_data)
            self.insert_text(f"拆分发票保存完毕", "green")
            QApplication.processEvents()  # 处理事件队列

            QMessageBox.information(self, "成功", "处理完成")
        except Exception as e:
            self.insert_text(f"处理发票出错, 错误原因:\n{e}", "red")
            QApplication.processEvents()  # 处理事件队列

            # error message box
            QMessageBox.critical(self, "错误", "处理发票出错, 请查看错误信息")

            print(e)

    def save_excel(self, empty_info_df: pd.DataFrame, service_template_df: pd.DataFrame):
        emtpy_info_df_save_path = os.path.join(self.save_dir_path, "未找到信息的发票.xlsx")
        service_template_df_save_path = os.path.join(self.save_dir_path, "service_template.xlsx")
        empty_info_df.to_excel(emtpy_info_df_save_path, index=False)
        service_template_df.to_excel(service_template_df_save_path, index=False)




    def open_xlsx_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "打开 Excel 文件", "",
                                                   "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if file_name:
            if file_name.endswith(('.xlsx', '.xls')):  # 确保包含所有格式的元组
                print(f"打开文件: {file_name}")
                # TODO: Load the video file into the video player widget
                # Show success message box
                QMessageBox.information(self, "成功", f"成功打开文件：{file_name}")
                return file_name
            else:
                QMessageBox.warning(self, "错误", "文件格式不支持")
                return None
        else:
            # Show error message box
            QMessageBox.critical(self, "错误", "未选择文件或文件打开失败")
            return None

    def handle_video_init(self):
        print("视频处理初始化函数")
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "打开视频文件", "", "All Files (*);;Video Files (*.mp4 *.avi)",
                                                   options=options)
        if file_name:
            if file_name.endswith(('.mp4', '.avi')):
                print(f"打开文件: {file_name}")
                # TODO: Load the video file into the video player widget
                self.video_path = file_name

                # Show success message box
                QMessageBox.information(self, "成功", f"成功打开文件：{file_name}")

                # 加载模型
                self.model = load_model()
                # self.show_progress_dialog()
                QMessageBox.information(self, "提示", "模型加载完成")
            else:
                QMessageBox.warning(self, "错误", "文件格式不支持")
        else:
            # Show error message box
            QMessageBox.critical(self, "错误", "未选择文件或文件打开失败")

    def handle_video_process(self):
        # 尝试访问self.video_path
        try:
            if self.model:
                if self.video_path:
                    self.play_video(self.video_path)
                    # self.
                else:
                    QMessageBox.warning(self, "错误", "未选择文件或文件打开失败, 请重新选择文件")
            else:
                QMessageBox.warning(self, "错误", "未加载模型,请先加载模型")
        except:
            QMessageBox.warning(self, "错误", "请先打开视频文件")

            return

    def handle_video(self):
        print("视频处理函数")
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "打开视频文件", "", "All Files (*);;Video Files (*.mp4 *.avi)",
                                                   options=options)
        if file_name:
            if file_name.endswith(('.mp4', '.avi')):
                print(f"打开文件: {file_name}")
                # TODO: Load the video file into the video player widget
                self.video_path = file_name

                # Show success message box
                QMessageBox.information(self, "成功", f"成功打开文件：{file_name}")

                # 加载模型
                self.model = load_model()
                # self.show_progress_dialog()
                QMessageBox.information(self, "提示", "模型加载完成")
                self.play_video(self.video_path)
            else:
                QMessageBox.warning(self, "错误", "文件格式不支持")
        else:
            # Show error message box
            QMessageBox.critical(self, "错误", "未选择文件或文件打开失败")

    def play_video(self, file_name):
        cap = cv2.VideoCapture(file_name)
        self.last_time = 0
        while cap.isOpened() and self.running:
            ret, frame = cap.read()
            if not ret:
                break
            detect_image = detect_box(self.model, frame)
            image_to_display = detect_image  # np.hstack((frame, detect_image))
            print(f"image to display shape {image_to_display.shape}")
            # Convert the frame to QImage and display it in the output_image_label
            height, width, channel = image_to_display.shape
            bytes_per_line = channel * width
            qimage = QImage(image_to_display.data, width, height, bytes_per_line, QImage.Format_BGR888)
            self.output_image_label.setPixmap(QPixmap.fromImage(qimage))
            self.output_image_label.setScaledContents(True)
            # Add a delay to control the video playback speed
            # time.sleep(0.02)
            # Process events to update the GUI
            QCoreApplication.processEvents()
            self.frame_has_been_detected_count += 1
            self.frame_count.setText(f"已经检测帧数：{self.frame_has_been_detected_count}")
            self.current_time = time.time()
            # 如果上一次检测到的帧的时间不为0，计算每一帧的运行时间和FPS
            if self.last_time != 0:
                frame_time = self.current_time - self.last_time
                fps = 1 / frame_time
                # 打印FPS
                self.fps.setText(f"FPS: %.2f" % fps)
            # 更新上一次检测到的帧的时间为当前的时间
            self.last_time = self.current_time

        # Release the video capture object
        cap.release()

    def play_udp(self):
        # 取出deque中的最后一帧
        image = self.image_deque.pop()
        self.frame_has_been_detected_count += 1
        self.frame_count.setText(f"已经检测帧数：{self.frame_has_been_detected_count}")
        self.current_time = time.time()
        # 如果上一次检测到的帧的时间不为0，计算每一帧的运行时间和FPS
        if self.last_time != 0:
            frame_time = self.current_time - self.last_time
            fps = 1 / frame_time
            # 打印FPS
            self.fps.setText(f"FPS: %.2f" % fps)
        # 更新上一次检测到的帧的时间为当前的时间
        self.last_time = self.current_time
        detect_image = detect_box(self.model, image)
        image_to_display = detect_image  # np.hstack((image, detect_image))
        # Convert the frame to QImage and display it in the output_image_label
        height, width, channel = image_to_display.shape
        bytes_per_line = channel * width
        qimage = QImage(image_to_display.data, width, height, bytes_per_line, QImage.Format_BGR888)
        self.output_image_label.setPixmap(QPixmap.fromImage(qimage))
        self.output_image_label.setScaledContents(True)
        # Add a delay to control the video playback speed
        # time.sleep(0.01)
        # Process events to update the GUI
        QCoreApplication.processEvents()

    def handle_camera(self):
        print("摄像头处理函数")

    def handle_pcie(self):
        print("PCIE处理函数")

    def handle_udp_init(self):
        print("UDP处理初始化函数")
        try:
            # TODO, 将生产者，消费者模型加入
            self.image_socket = ImageSocket()
            QMessageBox.information(self, "成功", f"成功初始化UDP")
            # self.image_processor = ImageProcessor()
            # QMessageBox.information(self, "成功", f"成功初始化UDP线程池")

        except Exception as e:
            QMessageBox.warning(self, "错误", "初始化UDP失败\n 失败原因：" + str(e))
        self.model = load_model()
        QMessageBox.information(self, "提示", "模型加载完成")

    def handle_udp_process(self):
        print("UDP处理函数")
        try:
            if self.model:
                if self.image_socket:
                    while self.running:
                        try:
                            # # 第一部分
                            # self.image_socket.send(hex_data="72656e54")
                            # buffer = self.image_socket.recv(1920 * 1080 * 4 // 1920 * 2, multi_thread=True)
                            # image = self.image_socket.buffer2image(1920, buffer)
                            # image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
                            # # image = udp_image_correct(image)
                            # self.image_socket.s.close()
                            # self.image_socket = ImageSocket()
                            # self.image_deque.append(image)
                            #
                            # # 第二部分
                            # self.play_udp()
                            # 第一部分
                            start_time = time.time()  # 记录开始时间
                            self.image_socket.send(hex_data="72656e54")
                            print(f"send_hex_data Time: {time.time() - start_time} seconds")

                            start_time = time.time()  # 重设开始时间
                            buffer = self.image_socket.recv(1920 * 1080 * 4 // (1920 * 4), multi_thread=True)
                            print(f"recv Time: {time.time() - start_time} seconds")

                            start_time = time.time()  # 重设开始时间
                            image = self.image_socket.buffer2image(1920, buffer)
                            print(f"buffer2image Time: {time.time() - start_time} seconds")

                            start_time = time.time()  # 重设开始时间
                            image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
                            print(f"cvtColor Time: {time.time() - start_time} seconds")

                            start_time = time.time()  # 重设开始时间
                            self.image_socket.s.close()
                            self.image_socket = ImageSocket()
                            print(f"renew Socket Time: {time.time() - start_time} seconds")

                            start_time = time.time()  # 重设开始时间
                            self.image_deque.append(image)
                            print(f"append Image Time: {time.time() - start_time} seconds")

                            # 第二部分
                            start_time = time.time()  # 重设开始时间
                            self.play_udp()
                            print(f"play_udp Time: {time.time() - start_time} seconds")



                        except Exception as e:
                            print(f"UDP 处理出错{e}")
                else:
                    QMessageBox.warning(self, "错误", "请先初始化UDP")
            else:
                QMessageBox.warning(self, "错误", "请先加载模型")
        except Exception as e:
            QMessageBox.warning(self, "错误", "请先初始化UDP\n 失败原因：" + str(e))
            return

    def handle_udp(self):
        print("UDP处理函数")
        self.image_socket = ImageSocket()
        QMessageBox.information(self, "成功", f"成功初始化UDP")
        self.model = load_model()
        QMessageBox.information(self, "提示", "模型加载完成")
        while True:
            try:
                print(f"开始发送数据")
                start_time = time.time()
                self.image_socket.send(hex_data="72656e54")
                end_time = time.time()
                print(f"发送数据耗时: {end_time - start_time} 秒")
                start_time = time.time()
                print(f"开始接受数据")
                buffer = self.image_socket.recv(1920 * 1080 * 4 // 1920 * 2, multi_thread=True)
                end_time = time.time()
                print(f"接受数据耗时: {end_time - start_time} 秒")
                start_time = time.time()
                print(f"开始数据转换成图片")
                image = self.image_socket.buffer2image(1920, buffer)
                end_time = time.time()
                print(f"数据转换耗时: {end_time - start_time} 秒")
                if (image.shape[0] == 0):
                    print(image.shape)
                    start_time = time.time()
                    self.image_socket.send(hex_data="72656e54")
                    buffer = self.image_socket.recv(1920 * 1080 // 1920)
                    image = self.image_socket.buffer2image(1920, buffer)
                    end_time = time.time()
                    print(f"重新获取和转换数据耗时: {end_time - start_time} 秒")
                start_time = time.time()
                image = image / 255
                plt.imsave("image.png", image)
                image = cv2.imread("image.png")
                image = udp_image_correct(image)
                end_time = time.time()
                print(f"图像处理耗时: {end_time - start_time} 秒")
                print(image.shape)
                print(f"开始关闭链接")
                start_time = time.time()
                self.image_socket.s.close()
                end_time = time.time()
                print(f"关闭链接耗时: {end_time - start_time} 秒")
                print(f"开始重置链接")
                start_time = time.time()
                self.image_socket = ImageSocket()
                end_time = time.time()
                print(f"重置链接耗时: {end_time - start_time} 秒")
                print(f"开始检测图片并显示")
                start_time = time.time()
                self.play_udp(image)
                end_time = time.time()
                print(f"检测图片并显示耗时: {end_time - start_time} 秒")

            except Exception as e:
                print(f"UDP 处理出错{e}")




