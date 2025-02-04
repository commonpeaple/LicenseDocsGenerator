import TkEasyGUI as sg
import pandas as pd
import subprocess
import shutil
from enum import IntEnum
from typing import List
from typing import overload

#=========前回選択した保存先を読み出す=========
save_path_file_path = 'saved_path.txt'
save_path_num = 4
class SavedPathIndex(IntEnum):
    TABLE=0
    EXDOC=1
    INDOC=2
    BUILD=3

line_count = 0
saved_path_list = []
try:
    with open(save_path_file_path,mode='r', encoding='utf-8') as f:
        for s_line in f:
            saved_path_list.append(s_line[:-1])
            line_count+=1
    if line_count <save_path_num:
        for i in range(save_path_num-line_count):
            saved_path_list.append('')
        line_count=0
        
            
except FileNotFoundError:
    for i in range(save_path_num):
        saved_path_list.append('')

#=========関数定義=========
#エクセルに記入
def write_excel(write_df,table_path):
    try:  
        with pd.ExcelWriter(table_path,mode='a',if_sheet_exists="replace") as writer:
                write_df =write_df.to_excel(writer,index=False)
    except PermissionError as e:
        sg.popup(f"ファイル \"{table_path}\" を閉じて下さい")
    
#ゲーム外ドキュメント作成
def write_exdox(license_table_df):
    #=====ゲーム外ドキュメント作成=====
        prev_category=''
        ex_doc_contents ='〇以下は使用したツール・アセットのライセンス文章です\n\n'
        for index,row in license_table_df.iterrows():
            license_url_flag = False
            if row[license_path_key]!=float('nan'):
                if 'http' in str(row[license_path_key]):
                    license_url_flag = True
            
            third_party_notice_flag = False
            if str(row[third_party_notice_document_key])!=str('nan'):
                third_party_notice_flag=True
            
            ex_doc_contents +=\
                f'{'---------------------------------------\n'\
                f'{row[category_key]}\n'\
                f'---------------------------------------\n' if prev_category!=str(row[category_key]) else ''}'\
                f'●{row[asset_name_key]}\n'\
                f'Asset Name: {row[asset_name_key]}\n'\
                f'Holder/Copyright: {row[copyright_holder_name_key]}\n'\
                f'URL: {row[download_url_key]}\n'\
                f'{f'License URL: {row[license_path_key]}\n' if (license_url_flag) else ''}'\
                f'License Type: {row[license_category_key]}\n'\
                f'License Document:\n'\
                f'{row[license_document_key]}\n\n'\
                f'{f'●ThirdPartyNotice of {row[asset_name_key]}\n'\
                f'{row[third_party_notice_document_key]}\n' if third_party_notice_flag else ''}\n'
            
            if str(row[category_key])==str('nan'):
                prev_category = str('nan')
            else:
                prev_category = row[category_key]
        try:
            with open(input_dict[ex_document_path_key],mode='w', encoding='utf-8') as f:
                f.write(ex_doc_contents)
        except FileNotFoundError:
            sg.popup(f"パス \"{ex_doc_contents}\" は存在していません")

#ゲーム内ドキュメント作成
def write_indox(license_table_df):
    prev_category=''
    in_doc_contents = '●クレジット・ライセンス\n'
    for index,row in license_table_df.iterrows():
        in_doc_contents +=\
            f'{f'〇{row[category_key]}\n' if prev_category!=str(row[category_key]) else ''}'\
            f'・{row[asset_name_key]}'\
            f'/{row[copyright_holder_name_key]}\n\n'
        if str(row[category_key])==str('nan'):
            prev_category = str('nan')
        else:
            prev_category = row[category_key]
    try:  
        with open(input_dict[in_document_path_key],mode='w', encoding="utf-8") as f:
            f.write(in_doc_contents)
    except FileNotFoundError:
        sg.popup(f"パス \"{in_doc_contents}\" は存在していません")

    

#パスを保存する
@overload
def write_save_path(save_path_list:List[str],change_index:int,value:str)->None:
    
    try:
        with open(save_path_file_path,mode='w', encoding="utf-8") as f:
            f.write('\n'.join(save_path_list))
    except FileNotFoundError:
        sg.popup(f"パス \"{save_path_file_path}\" は存在していません")

@overload
def write_save_path(save_path_list:List[str])->None:
    pass

def write_save_path(save_path_list,change_index=None,value=None)->None:
    if change_index != None: 
        save_path_list[change_index] = value
    try:
        with open(save_path_file_path,mode='w', encoding="utf-8") as f:
            f.write('\n'.join(save_path_list))
    except FileNotFoundError:
        sg.popup(f"パス \"{save_path_file_path}\" は存在していません")


#=========変数定義=========
#入力部分のサイズ
size_param = [40,1]
#複数行入力のサイズ
multiline_size_param = [40,3]
#未入力情報が含まれるかを示すフラグ(これが無いと、未入力時落ちる)
nan_flag = False

#使用用途情報
category_key = "category"
#アセット情報
asset_name_key = "asset_name"
copyright_holder_name_key="copyright_holder_name"
download_url_key="download_url"
#ライセンス情報
license_category_key="license_category"
license_document_key="license_document"
license_path_key="license_path"
#ThirdPartyNotice情報
third_party_notice_document_key="third_party_notice_document"
#ドキュメント保存情報
license_table_path_key = "license_table_path"
ex_document_path_key ="external_license_document_path"
in_document_path_key ="internal_license_document_path"
    
#ライセンステーブルの列名
license_table_columns=[
    #使用用途情報
    category_key,
    "usage_purpose",
    "asset_path",
    
    #アセット情報
    asset_name_key,
    copyright_holder_name_key,
    download_url_key,

    #ライセンス情報
    license_category_key,
    license_document_key,
    license_path_key,
    
    #ThirdPartyNotice情報
    "third_party_notice_category",
    third_party_notice_document_key,
    "third_party_notice_path",
        
    #ドキュメント保存情報
    license_table_path_key,
    ex_document_path_key,
    in_document_path_key,
    "additive_info",
]

#ゲーム外ライセンスの要素
external_license=[
    "category",
    
    #アセット情報
    "asset_name",
    "copyright_holder_name",
    "download_url",

    #ライセンス情報
    "license_category",
    "license_document",
    "license_path",
    
    #ThirdPartyNotice情報
    "third_party_notice_category",
    "third_party_notice_document",
    "third_party_notice_path",
]

#ゲーム内ライセンスの要素
internal_license_columns=[
    "category",
    
    #アセット情報
    "asset_name",
    "copyright_holder_name",
    "download_url",
]


#==========レイアウト定義==========
#======アセット情報ブロック=====
asset_layout=[
    [sg.Text(
            "<アセット情報>" # ラベル
            ,color="white"
            ,background_color="black"
    )],
    #===アセット名===
    [sg.Text(
            "アセット名(必須):" # ラベル
    )],
    [sg.Input("",
            key="-asset_name-" # 要素の参照キー
            ,size=size_param
        ), 
    ],
    
    #===著作者名===
    [sg.Text(
            "著作者名(必須):" # ラベル
    )],
    [sg.Input("",
            key="-copyright_holder_name-" # 要素の参照キー
            ,size=size_param
        ), 
    ],
    
    #===ダウンロードURL===
    [sg.Text(
            "ダウンロードURL(必須):" # ラベル
    )],
    [sg.Input("",
            key="-download_url-" # 要素の参照キー
            ,size=size_param
        ), 
    ]
]

#======ライセンス情報ブロック=====
license_layout=[
    [sg.Text(
            "<ライセンス情報>" # ラベル
            ,color="white"
            ,background_color="black"
    )],
    #===ライセンス種類===
    [sg.Text(
            "ライセンス種類(必須):" # ラベル
    )],
    [sg.Input("",
            key="-license_category-" # 要素の参照キー
            ,size=size_param
        ), 
    ],
    
    #===ライセンス文章===
    [sg.Text(
            "ライセンス文章(必須):" # ラベル
    )],
    [sg.Multiline("",
            key="-license_document-" # 要素の参照キー
            ,size=multiline_size_param
        ), 
    ],
    #===ライセンスパス/URL===
    [sg.Text(
            "ライセンスパス(必須):" # ラベル
    )],
    [sg.Input("",
            key="-license_path-" # 要素の参照キー
            ,size=size_param
        ), 
        sg.FileBrowse()
    ]
]
    
#=====ThirdPartyNoticeパス=====
third_party_notice_layout=[
    [sg.Text(
            "<ThirdPartyNotice情報>" # ラベル
            ,color="white"
            ,background_color="black"
    )],
    #ThirdPartyNotice種類
    [sg.Text(
            "ThirdPartyNotice種類:" # ラベル
    )],
    [sg.Multiline("",
            key="-third_party_notice_category-" # 要素の参照キー
            ,size=multiline_size_param
        ), 
    ],
    #ThirdPartyNotice文章
    [sg.Text(
            "ThirdPartyNotice文章:" # ラベル
    )],
    [sg.Multiline("",
            key="-third_party_notice_document-" # 要素の参照キー
            ,size=multiline_size_param
        ), 
    ],
    #ThirdPartyNoticeパス
    [sg.Text(
            "ThirdPartyNoticeパス:" # ラベル
    )],
    [sg.Multiline("",
            key="-third_party_notice_path-" # 要素の参照キー
            ,size=multiline_size_param
        ), 
        sg.FileBrowse()
    ]
]

#=====使用用途情報=====
usage_layout = [
    [sg.Text(
            "<使用用途情報>" # ラベル
            ,color="white"
            ,background_color="black"
    )],
    #===カテゴリ===
    [sg.Text(
            "カテゴリ(必須):" # ラベル
    )],
    [sg.Input(
            "", # テキスト
            key="-category-" ,
            size=size_param
        ),
    ], 
    #===使用用途===
    [sg.Text(
            "使用用途(必須):" # ラベル
    )],
    [sg.Input("",
            key="-usage_purpose-", # 要素の参照キー
            size=size_param
        ), 
    ],
    #===アセットパス===
    [sg.Text(
        "アセットパス(必須):" # ラベル
    )],
    [sg.Input(
            "", # テキスト
            key="-asset_path-" # 要素の参照キー
            ,size=size_param
            ,expand_x=True
            ,expand_y=True
        ), 
        sg.FileBrowse() # 単一ファイル選択
    ]
]
    
#=====ドキュメント保存情報=====
document_path_layout =[
    [sg.Text(
            "<ドキュメント保存情報>" # ラベル
            ,color="white"
            ,background_color="black"
    )],
    #===ライセンステーブルパス===
    [sg.Text(
        "ライセンステーブルパス(必須):" # ラベル
    )],
    [sg.Input(
            saved_path_list[SavedPathIndex.TABLE], # テキスト
            key="-license_table_path-" # 要素の参照キー
            ,size=[50,100]
            ,expand_x=True
            ,expand_y=True
        ), 
        sg.FileBrowse() # 単一ファイル選択
    ], 
    #====ライセンスドキュメント(ゲーム外)===
    [sg.Text(
        "ライセンスドキュメントパス(ゲーム外)(必須):" # ラベル
    )],
    [sg.Input(
            saved_path_list[SavedPathIndex.EXDOC], # テキスト
            key="-external_license_document_path-" # 要素の参照キー
            ,size=[30,100]
            ,expand_x=True
            ,expand_y=True
        ), 
        sg.FileBrowse() # 単一ファイル選択
    ], 
    #====ライセンスドキュメント(ゲーム内)===
    [sg.Text(
        "ライセンスドキュメントパス(ゲーム内)(必須):" # ラベル
    )],
    [sg.Input(
            saved_path_list[SavedPathIndex.INDOC], # テキスト
            key="-internal_license_document_path-" # 要素の参照キー
            ,size=[50,100]
            ,expand_x=True
            ,expand_y=True
        ), 
        sg.FileBrowse() # 単一ファイル選択
    ]
]

#=====備考/参考文献=====
additive_info_layout = [
    #===備考/参考文献===
    [sg.Text(
            "<備考/参考文献>" # ラベル
            ,color="white"
            ,background_color="black"
    )],
    [sg.Multiline("",
            key="-additive_info-" # 要素の参照キー
            ,size=size_param
        ), 
    ]
    
]

#=====ボタンレイアウト=====
button_layout = [   
    [sg.Button(
        "Save", # ラベル
    ),
     sg.Checkbox(
        "Save後Clearするか", #ラベル
        key = "-checkbox-",
        enable_events=True
    ),
    sg.Button(
        "Clear" # ラベル
    )],
    
    [sg.Button(
        "SavePath" #ラベル   
    ),
     sg.Button(
        "OpenDocs" # ラベル
    ),
    sg.Button(
        "SyncExcel" # ラベル
    )],
    #====ビルドに付属するときに使う===
    [sg.Text(
        "ビルド/プロジェクトフォルダ:" # ラベル
    )],
    [sg.Input(
            saved_path_list[SavedPathIndex.BUILD], # テキスト
            key="-build_folder-" # 要素の参照キー
            ,size=[50,100]
            ,expand_x=True
            ,expand_y=True
        ), 
        sg.FolderBrowse() 
    ],
    [sg.Button(
        "PlaceEXDocs" # ラベル
    )],
]

#==========レイアウト組み合わせ==========
left_layout = asset_layout+[[sg.HSeparator()]]+license_layout+[[sg.HSeparator()]]+third_party_notice_layout

right_layout = usage_layout+[[sg.HSeparator()]]+document_path_layout+[[sg.HSeparator()]]+button_layout+[[sg.HSeparator()]]+additive_info_layout

col1=sg.Column(left_layout,key="col1",expand_x=True,expand_y=True)
col2=sg.Column(right_layout,key="col2")
layout = [
    [col1,sg.VSeparator(),col2]
]

# ウィンドウのレイアウトを定義
window = sg.Window("License Document Generator", resizable=True,layout=layout)

#==========イベントループ==========
while window.is_alive():
    nan_flag=False
    event, values = window.read()
    #===保存処理===
    if event == "Save":
        input_dict={}
        #=====入力の取得=====
        #入力をディクショナリで保存
        for key in license_table_columns:
            if (values[f"-{key}-"] == '')and (key in license_table_columns[0:9] or key in license_table_columns[13:15]):
                sg.popup(f'必須項目を入力して下さい')
                nan_flag=True
                break
            else:
                input_dict[key]=values[f"-{key}-"]
        if nan_flag:
            continue
        
        #入力を1行データフレームに変換
        new_asset_row_df = pd.DataFrame([[values[f"-{key}-"] for key in license_table_columns]],columns=license_table_columns)
        new_asset_row_df = new_asset_row_df.replace('',float('nan'))
        
        #=====エクセル処理=====
        #エクセルデータ取得
        try:
            license_table_df = pd.read_excel(input_dict[license_table_path_key])
        except FileNotFoundError:
            sg.popup(f"パス \"{license_table_path_key}\" は存在していません")
    
        #既にアセットが登録されているかチェック
        try:
            already_asset_index=license_table_df.index[license_table_df[asset_name_key]==input_dict[asset_name_key]]
        except KeyError:
            already_asset_index=None
            df = pd.DataFrame([],columns=license_table_columns)
            write_excel(df,input_dict[license_table_path_key])

        #既にアセットが登録されている時
        if already_asset_index!=None:
            res = sg.popup_yes_no(f"アセット \"{input_dict[asset_name_key]}\" は既に登録されています。上書きしますか?")
            if res=="Yes":
                license_table_df.iloc[already_asset_index] =  new_asset_row_df
                #出力
                license_table_df = license_table_df.sort_values(category_key,na_position='last')
                write_excel(license_table_df,input_dict[license_table_path_key])
            else:
                continue
        
        #アセット新規登録
        else:
            #既存データと結合して、カテゴリでソート
            license_table_df = pd.concat([license_table_df,new_asset_row_df],axis=0)
            license_table_df = license_table_df.sort_values(category_key,na_position='last')
            #出力
            write_excel(license_table_df,input_dict[license_table_path_key])
             
        #=====ゲーム外ドキュメント作成=====
        write_exdox(license_table_df)
            
        #=====ゲーム内ドキュメント作成=====
        write_indox(license_table_df)

        #=====ドキュメント保存情報の保存=====
        write_save_path(
            [values[f"-{license_table_path_key}-"],
             values[f"-{ex_document_path_key}-"],
             values[f"-{in_document_path_key}-"],
             saved_path_list[SavedPathIndex.BUILD]]
            )
        
        sg.popup('正常に保存されました')  
        
        if  window["-checkbox-"].get():
            for key in license_table_columns[0:9]:
                window[f"-{key}-"].update('')        
            window[f"-{license_table_columns[-1]}-"].update('')
        continue
    
    #===入力部分を削除===
    if event ==  "Clear":
        for key in license_table_columns[0:9]:
            window[f"-{key}-"].update('')        
        window[f"-{license_table_columns[-1]}-"].update('')
    
    #===ドキュメント開く===
    if event == "OpenDocs":
        popen =subprocess.Popen(["start","",values[f"-{license_table_path_key}-"]],shell=True)
        popen.wait()
        popen =subprocess.Popen(["start","",values[f"-{ex_document_path_key}-"]],shell=True)
        popen.wait()
        popen =subprocess.Popen(["start","",values[f"-{in_document_path_key}-"]],shell=True)
        popen.wait()
        
        #=====ドキュメント保存情報を保存=====
        write_save_path([
            values[f"-{license_table_path_key}-"],
            values[f"-{ex_document_path_key}-"],
            values[f"-{in_document_path_key}-"],
            saved_path_list[SavedPathIndex.BUILD]
        ])
        
        sg.popup(f'ライセンス表,ゲーム外/ゲーム内ドキュメントを開きました\\また、ドキュメント保存情報を保存しました')    
    
    #===ドキュメント保存情報とビルドフォルダパスを保存===
    if event == "SavePath":
        
        write_save_path([
            values[f"-{license_table_path_key}-"],
            values[f"-{ex_document_path_key}-"],
            values[f"-{in_document_path_key}-"],
            values[f"-{ex_document_path_key}-"]
        ])
        sg.popup(f'ドキュメント保存情報とビルドフォルダのパスが保存されました。') 
    
    #===ドキュメントを置く(ビルドフォルダが上書きされて外部ドキュメントが消えるため用意)===
    if event == "PlaceEXDocs":
        try:
            shutil.copy(values[f"-{ex_document_path_key}-"],values['-build_folder-'], follow_symlinks=True)
            write_save_path(saved_path_list,SavedPathIndex.BUILD,values[f"-{ex_document_path_key}-"])
            sg.popup(f'指定パスにコピーしました\\また、ビルドフォルダパスを保存しました')
            
        except shutil.SameFileError:
            r= sg.popup_ok_cancel("指定パスに同一ファイルがあります。\n上書きしますか?")
            if r=='OK':
                shutil.rmtree(values['-build_folder-']+'\\'+values[f"-{ex_document_path_key}-"])
                shutil.copy(values[f"-{ex_document_path_key}-"],values['-build_folder-'], follow_symlinks=False)
                write_save_path(saved_path_list,SavedPathIndex.BUILD,values[f"-{ex_document_path_key}-"])
                sg.popup(f'指定パスにコピーしました\\また、ビルドフォルダパスを保存しました')
        except FileNotFoundError:
            sg.popup(f'ビルドフォルダパスを入力してください')

    #===Excelと同期させる===
    if event == "SyncExcel":
        input_dict={}
        #=====入力の取得=====
        #入力をディクショナリで保存
        for key in license_table_columns:
            if (values[f"-{key}-"] == '')and  key in license_table_columns[13:15]:
                sg.popup(f'ドキュメント保存情報を入力してください')
                nan_flag=True
                break
            else:
                input_dict[key]=values[f"-{key}-"]
        if nan_flag:
            continue
        
        #エクセルデータ取得
        try:
            license_table_df = pd.read_excel(input_dict[license_table_path_key])
        except FileNotFoundError:
            sg.popup(f"パス \"{license_table_path_key}\" は存在していません")

        #=====ゲーム外ドキュメント作成=====
        write_exdox(license_table_df)
            
        #=====ゲーム内ドキュメント作成=====
        write_indox(license_table_df)
        
        write_save_path([values[f"-{license_table_path_key}-"],values[f"-{ex_document_path_key}-"],values[f"-{in_document_path_key}-"],saved_path_list[SavedPathIndex.BUILD]])
        sg.popup("エクセルファイルの内容で、ライセンスファイルを更新しました") 

    #最大化/最小化
    if event == "Maximize":
        # ウィンドウの最大化
        window.maximize()
    if event == "Minimize":
        # ウィンドウの縮小化
        window.minimize()
    
window.close()