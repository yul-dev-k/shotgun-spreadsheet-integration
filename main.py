import gspread
from oauth2client.service_account import ServiceAccountCredentials
import shotgun_api3
import os
from dotenv import load_dotenv
import time
from tkinter import *
from tkinter import ttk, messagebox

load_dotenv()

URL = os.environ.get("BASE_URL")
LOGIN = os.environ.get("LOGIN")
PW = os.environ.get("PASSWORD")

sg = shotgun_api3.Shotgun(URL, login=LOGIN, password=PW)

scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
]

# google cloud에서 발급받은 JSON형태의 api 파일이 필요합니다.
json_file_name = './sheetapi.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    json_file_name, scope)
gc = gspread.authorize(credentials)
spreadsheet_url = os.environ.get("SPREADSHEET_URL")
doc = gc.open_by_url(spreadsheet_url)


class WorkSheet:
    start_row = 24
    code_col, ani_b_col, bg_col, amb_col, assign_col, bg_comment_col = [
        2, 3, 4, 5, 6, 13]
    retry_count = 0
    max_retries = 20

    def __init__(self, title):
        self.title = title

    def get_shot_data(self, episode_id):
        """episode_id를 기반으로 shotgun_api3를 이용해 시트에 필요한 샷데이터를 받아옵니다.

        Args:
            episode_id (number): 샷건에 등록된 episode_id가 필요합니다.

        Returns:
            list: [{'type': 'Shot', 'id': , 'sg_uptoftp': '', 'code': '', 'sg_assets_bg': [{'id': , 'name': '', 'type': 'Asset'}], 'sg_ambience': '', 'sg_assigned': [{'id': , 'name': '', 'type': 'HumanUser'}], 'sg_bg_comments': None}]
        """
        return sorted(sg.find("Shot", [["sg_episode", "is", {"type": 'Episode', "id": episode_id}]],
                              ["sg_uptoftp", "code", "sg_assets_bg", "sg_ambience", "sg_assigned", "sg_bg_comments"]),
                      key=lambda x: x['code'])

    def calculate_dynamic_sleep(self, api_error):
        """google cloud API의 할당량을 초과했을 경우 60초 뒤에 실행하게 하는 예외 처리 함수입니다.

        Args:
            api_error (string): gspread.exceptions.APIError

        Returns:
            number: time.sleep을 얼마나 지정할지를 반환합니다.
        """
        retry_after_header = api_error.response.headers.get('Retry-After')
        if retry_after_header:
            try:
                return int(retry_after_header)
            except ValueError:
                pass
        return 60

    def initialize_worksheet(self, episode_id, progressbar, root):
        """get_shot_data()에서 받아온 데이터를 시트에 생성해주는 초기 함수입니다.

        Args:
            episode_id (number): 샷건에 등록된 episode_id가 필요합니다.
        """
        total_shots = len(self.get_shot_data(episode_id))
        if total_shots == 0:
            messagebox.showerror("ERROR", "shot 데이터가 없습니다.")
            return
        template = doc.worksheet("EP600")
        dup_sheet = template.duplicate(new_sheet_name=self.title)
        dup_sheet.update_cell(2, 2, value=self.title)
        dup_sheet.update_cell(
            4, 7, value=f"▼{self.title} mg 디자인 리스트▼\nX:/p38_MiniForce/01_preproduction/season06/03_design/01_final/{str(self.title).lower()}/\n▼ 키컷경로 ▼\nZ:/p38_MiniForce/01_COMP/season06/00_KEYCUT/")
        cell_list = []

        progressbar_step = 100 / total_shots

        for shot_index, shot in enumerate(self.get_shot_data(episode_id), start=self.start_row):
            print(f"현재 {shot['code']} 추가 중")
            cell_list.extend([
                gspread.cell.Cell(shot_index, self.code_col, shot["code"]),
                gspread.cell.Cell(shot_index, self.bg_col, shot["sg_assets_bg"][0]["name"] if len(
                    shot["sg_assets_bg"]) == 1 else None),
                gspread.cell.Cell(shot_index, self.amb_col,
                                  shot["sg_ambience"]),
                gspread.cell.Cell(shot_index, self.assign_col, (shot["sg_assigned"][0]["name"] if len(
                    shot["sg_assigned"]) == 1 else "재사용" if shot["sg_uptoftp"] == "reuse" else None)),
                gspread.cell.Cell(
                    shot_index, self.bg_comment_col, shot["sg_bg_comments"]),
                gspread.cell.Cell(shot_index, self.ani_b_col,
                                  True if shot["sg_uptoftp"] == "pub" else False)
            ])

            progressbar.step(progressbar_step)
            root.update_idletasks()

        try:
            dup_sheet.update_cells(cell_list)
        except gspread.exceptions.APIError as api_error:
            if 'quota' in str(api_error).lower() and self.retry_count < self.max_retries:
                sleep_time = self.calculate_dynamic_sleep(api_error)
                print(
                    f"API Quota Exceeded. Sleeping for {sleep_time} seconds.")
                time.sleep(sleep_time)
                self.retry_count += 1
                self.initialize_worksheet()
            else:
                raise
        print("시트 생성 완료")
        messagebox.showinfo("작업 완료", "시트 생성이 완료되었습니다.")

    def update_ani_backup(self, episode_id, progressbar, root):
        """ani backup된 것만 시트에 True처리 해줍니다.

        Args:
            episode_id (number): 샷건에 등록된 episode_id가 필요합니다.
        """
        worksheet = doc.worksheet(self.title)
        cell_list = []

        total_shots = len(self.get_shot_data(episode_id))
        progress_step = 100 / total_shots

        try:
            for shot_index, shot in enumerate(self.get_shot_data(episode_id), start=self.start_row):
                print(f"현재 {shot['code']} 업데이트 중")
                cell_list.append(gspread.cell.Cell(
                    shot_index, self.ani_b_col, value=True if shot["sg_uptoftp"] == "pub" else False))
                progressbar.step(progress_step)
                root.update_idletasks()

            worksheet.update_cells(cell_list)

        except gspread.exceptions.APIError as api_error:
            if 'quota' in str(api_error).lower() and self.retry_count < self.max_retries:
                sleep_time = self.calculate_dynamic_sleep(api_error)
                print(
                    f"API Quota Exceeded. Sleeping for {sleep_time} seconds.")
                time.sleep(sleep_time)
                self.retry_count += 1
                self.update_ani_backup()
            else:
                raise

        print("애니 백업 체크 완료")
        messagebox.showinfo("작업 완료", "애니 백업 업데이트가 완료되었습니다.")

    def duplicate_sheet(self):
        template = doc.worksheet("EP600")
        dup_sheet = template.duplicate(new_sheet_name=self.title)
        dup_sheet.update_cell(2, 2, value=self.title)


def widget():
    root = Tk()
    root.title("shotgun-googlesheet-연동-App")
    root.geometry("200x200+100+100")
    root.resizable(False, False)
    root['background'] = '#181914'
    root['padx'] = 15
    root['pady'] = 15

    episodes_id_name = [
        {'id': 4819, 'name': 'EP600'},
        {'id': 4820, 'name': 'EP601'},
        {'id': 4821, 'name': 'EP602'},
        {'id': 4822, 'name': 'EP603'},
        {'id': 4823, 'name': 'EP604'},
        {'id': 4824, 'name': 'EP605'},
        {'id': 4825, 'name': 'EP606'},
        {'id': 4826, 'name': 'EP607'},
        {'id': 4827, 'name': 'EP608'},
        {'id': 4828, 'name': 'EP609'},
        {'id': 4829, 'name': 'EP610'},
        {'id': 4830, 'name': 'EP611'},
        {'id': 4831, 'name': 'EP612'},
        {'id': 4832, 'name': 'EP613'},
        {'id': 4833, 'name': 'EP614'},
        {'id': 4834, 'name': 'EP615'},
        {'id': 4835, 'name': 'EP616'},
        {'id': 4836, 'name': 'EP617'},
        {'id': 4837, 'name': 'EP618'},
        {'id': 4838, 'name': 'EP619'},
        {'id': 4839, 'name': 'EP620'},
        {'id': 4840, 'name': 'EP621'},
        {'id': 4841, 'name': 'EP622'},
        {'id': 4842, 'name': 'EP623'},
        {'id': 4843, 'name': 'EP624'},
        {'id': 4844, 'name': 'EP625'},
        {'id': 4845, 'name': 'EP626'},
        {'id': 4846, 'name': 'EP627'},
        {'id': 4847, 'name': 'EP628'},
        {'id': 4848, 'name': 'EP629'},
    ]

    items = [name["name"] for name in episodes_id_name]

    left_frame = Frame(root, bg="#292929", height=60)
    left_frame.place(x=0, y=0)

    right_frame = Frame(root, bg="#181914", height=100, width=80)
    right_frame.place(x=60, y=0)

    listbox = Listbox(left_frame, height=10, width=6, selectmode="single")
    for item in items:
        listbox.insert(END, item)
    listbox.pack()
    listbox.selection_set(0)  # Set the default selection

    progressbar = ttk.Progressbar(right_frame, mode='determinate')
    progressbar.pack(fill=X, expand=True, pady=(0, 83))

    create_button = Button(right_frame, overrelief="solid", bg="#2E5EA2", fg="#ffffff", text="시트 생성", bd=0,
                           command=lambda: on_click_create_sheet(episodes_id_name, listbox.get(listbox.curselection()), progressbar, root))
    create_button.pack(fill=X, expand=True, pady=(0, 15))
    ani_up_button = Button(right_frame, overrelief="solid", bg="#2E5EA2",  fg="#ffffff", text="애니 백업 업데이트", bd=0,
                           command=lambda: on_click_ani_backup_update(episodes_id_name, listbox.get(listbox.curselection()), progressbar, root))
    ani_up_button.pack(fill=X, expand=True)

    root.mainloop()


def on_click_create_sheet(ep_list, selected_ep, progressbar, root):
    print(selected_ep)
    for selected_episode in ep_list:
        if selected_episode["name"] == selected_ep:
            worksheet_instance = WorkSheet(selected_episode["name"])
            progressbar.start()
            try:
                worksheet_instance.initialize_worksheet(
                    selected_episode["id"], progressbar, root)
            except gspread.exceptions.APIError as api_error:
                error_message = str(api_error)
                if 'duplicateSheet' in error_message and 'already exists' in error_message:
                    messagebox.showerror(
                        "ERROR", "이미 존재하는 시트입니다.")
                else:
                    raise
            progressbar.stop()


def on_click_ani_backup_update(ep_list, selected_ep, progressbar, root):
    print(selected_ep)
    for selected_episode in ep_list:
        if selected_episode["name"] == selected_ep:
            worksheet_instance = WorkSheet(selected_episode["name"])
            progressbar.start()
            worksheet_instance.update_ani_backup(
                selected_episode["id"], progressbar, root)
            progressbar.stop()


widget()
