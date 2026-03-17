import win32com.client
import subprocess
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import ssl

# ==========================================
# 1. 配置信息 (已脱敏)
# ==========================================
# 建议通过环境变量设置敏感信息，或建立本地 config.json
USER_INFO = {
    "user": os.getenv("SAP_USER", "YOUR_USER_HERE"), 
    "password": os.getenv("SAP_PASSWORD", "YOUR_PASSWORD_HERE")
}
SAP_SYSTEM_ID = os.getenv("SAP_SYSTEM_NAME", "SAP_SYSTEM_ALIAS_HERE")
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

# Google Sheet 配置
GS_KEY = os.getenv("GS_SHEET_KEY", "YOUR_GOOGLE_SHEET_ID_HERE")
CREDENTIALS_JSON = "credentials.json"  # 确保此文件已列入 .gitignore

# 业务常量提取
TARGET_PLANT = "CN12"
VIEW_NAME = "Quality Management"

try:
    ssl._create_default_https_context = ssl._create_unverified_context
except:
    pass


# ==========================================
# 2. Google Sheet 获取任务
# ==========================================
def get_mm01_tasks():
    print("\n📡 --- 步骤1: 获取任务 ---")
    try:
        # 权限范围
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        
        # 检查凭据文件是否存在
        if not os.path.exists(CREDENTIALS_JSON):
            print(f"错误: 找不到凭据文件 {CREDENTIALS_JSON}")
            return [], None, None

        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_JSON, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(GS_KEY)
        
        sheet_source = spreadsheet.worksheet("当日数据更新")
        try:
            sheet_done = spreadsheet.worksheet("已经维护的数据")
        except:
            sheet_done = None
            
        all_data = sheet_source.get_all_values()
        
        # 筛选逻辑：物料号以 P 开头且第 13 列不是 success
        tasks = []
        for i, row in enumerate(all_data):
            if i == 0: continue # 跳过表头
            
            p_no = str(row[1]).strip()
            # 这里的逻辑根据实际表格列数微调
            status = row[12].lower() if len(row) > 12 else ""
            
            if p_no.upper().startswith('P') and status != "success":
                tasks.append({"p_no": p_no, "row_idx": i + 1, "row_data": row})
                
        print(f" 待处理任务: {len(tasks)} 个")
        return tasks, sheet_source, sheet_done
    except Exception as e:
        print(f" Google Sheet 连接失败: {e}")
        return [], None, None


# ==========================================
# 3. SAP 登录与连接
# ==========================================
def get_sap_session():
    print("\n --- 步骤2: 登录 SAP ---")
    try:
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except:
            subprocess.Popen(SAP_LOGON_PATH)
            time.sleep(8)
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            
        application = SapGuiAuto.GetScriptingEngine
        connection = application.OpenConnection(SAP_SYSTEM_ID, True)
        session = connection.Children(0)
        
        # 登录界面处理
        if session.findById("wnd[0]/usr/txtRSYST-BNAME", False):
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = USER_INFO['user']
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = USER_INFO['password']
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(2)
            
            # 处理多重登录弹窗
            try:
                if "wnd[1]" in str(session.ActiveWindow.Name):
                    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except:
                pass
        print(" SAP 连接成功！")
        return session
    except Exception as e:
        print(f" SAP 连接失败: {e}")
        return None


def select_material_view(session, target_view=VIEW_NAME):
    """暴力遍历勾选视图"""
    print(f"       正在匹配视图: {target_view}...")
    try:
        # 重置全选
        try:
            session.findById("wnd[1]/usr/chkUSRM1-SISEL").selected = 0
        except:
            pass

        table = session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW")

        for i in range(25): # 增加探测范围
            try:
                view_text = table.getCell(i, 0).text
                if target_view.lower() in view_text.lower():
                    table.getAbsoluteRow(i).selected = True
                    time.sleep(0.5)
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    return True
            except:
                continue

        # 保底操作
        session.findById("wnd[1]").sendVKey(0)
        return True
    except Exception as e:
        print(f"       视图选择失败: {e}")
        return False


def run_sap_mm01(session, p_number):
    """执行 MM01 维护流程"""
    print(f"\n 处理物料: {p_number}")
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM01"
        session.findById("wnd[0]").sendVKey(0)
        
        session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = p_number
        session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "M"
        session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "VERP"
        session.findById("wnd[0]").sendVKey(0)

        # 1. 选择视图
        if not select_material_view(session):
            return False

        # 2. 填写工厂
        time.sleep(1)
        if "wnd[1]" in str(session.ActiveWindow.Name):
            session.activeWindow.findByName("RMMG1-WERKS", "GuiCTextField").text = TARGET_PLANT
            session.findById("wnd[1]").sendVKey(0)

        # 3. 循环回车进入主界面
        try:
            while not session.findById("wnd[0]/usr", False).findByName("MARC-QMATA", "GuiCTextField"):
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
        except:
            pass

        # 4. QM 维护逻辑
        qmata = session.findById("wnd[0]/usr").findByName("MARC-QMATA", "GuiCTextField")
        if qmata:
            qmata.text = "000024"
            session.findById("wnd[0]/usr").findByName("MARC_QMPD", "GuiButton").press()
            
            # 维护 Inspection Type 89 和 Z01
            btn_insert = session.findById("wnd[1]/tbar[0]/btn[5]")
            btn_insert.press()
            
            # 第一行 89
            session.findById("wnd[1]/usr/tblSAPLQPLSPRUEFDAT/ctxtRMQAM-ART[1,0]").text = "89"
            session.findById("wnd[1]/usr/tblSAPLQPLSPRUEFDAT/chkRMQAM-AKTIV[4,0]").selected = -1
            
            # 第二行 Z01
            session.findById("wnd[1]/usr/tblSAPLQPLSPRUEFDAT/ctxtRMQAM-ART[1,1]").text = "Z01"
            session.findById("wnd[1]/usr/tblSAPLQPLSPRUEFDAT/chkRMQAM-AKTIV[4,1]").selected = -1
            session.findById("wnd[1]/usr/tblSAPLQPLSPRUEFDAT/chkRMQAM-APA[3,1]").selected = -1

            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/btn[11]").press()  # 保存
            print(f"       物料 {p_number} 保存成功")
            return True
        return False

    except Exception as e:
        print(f"       流程出错: {e}")
        return False


if __name__ == "__main__":
    tasks, sheet_source, sheet_done = get_mm01_tasks()
    if tasks:
        sap_session = get_sap_session()
        if sap_session:
            for item in tasks:
                if run_sap_mm01(sap_session, item['p_no']):
                    try:
                        # 更新状态到第 13 列
                        sheet_source.update_cell(item['row_idx'], 13, "success")
                        if sheet_done:
                            log_data = item['row_data'].copy()
                            log_data.append(time.strftime("%Y-%m-%d %H:%M:%S"))
                            sheet_done.append_row(log_data)
                    except Exception as e:
                        print(f"⚠️ 更新 Google Sheet 状态失败: {e}")
    
    print("\n✨ 所有任务处理尝试完毕。")
    time.sleep(5)
