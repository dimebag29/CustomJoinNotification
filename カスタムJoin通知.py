# ==============================================================================================================
# 作成者:dimebag29 作成日:2023年11月19日 バージョン:v0.0
# (Author:dimebag29 Creation date:November 19, 2023 Version:v0.0)
#
# このプログラムのライセンスはCC0 (クリエイティブ・コモンズ・ゼロ)です。いかなる権利も保有しません。
# (This program is licensed under CC0 (Creative Commons Zero). No rights reserved.)
# https://creativecommons.org/publicdomain/zero/1.0/
#
# 開発環境 (Development environment)
# ･python 3.7.5
# ･auto-py-to-exe 2.36.0 (used to create the exe file)
#
# exe化時のauto-py-to-exeの設定
# ･ひとつのファイルにまとめる (--onefile)
# ･コンソールベース (--console)
# ･exeアイコン設定 (--icon)
# ==============================================================================================================

# python 3.7.5の標準ライブラリ (Libraries included as standard in python 3.7.5)
import os
import sys
import time
import glob
from itertools import islice
import datetime
import re
import socket
import json
import copy

# 外部ライブラリ (External libraries)
import win32gui                                                                 # Included in pywin32 Version:306
import psutil                                                                   # Version:5.9.6


# 関数定義
# ==============================================================================================================
# コンソールで文字に色を付けてprintする関数 参考:Pythonのprintいろいろ 文字色 https://qiita.com/yuto16/items/cb2f812b0d966b6ae920
def print_color(text, color="red"):
    print(color_dic[color] + text + color_dic["end"])


# VRChatのログフォルダ内から、名前の日付が一番新しいログファイルのパスを返す関数
def GetNewestLogFilePath():
    # LogFileDir内でファイル名検索文字列LogFileSearchWordにヒットしたファイルのパスリスト
    LogFilePathList = list(glob.glob(LogFileDir + LogFileSearchWord))

    # ひとつもログファイルを取得できなかったらエラーを出して終了
    if 0 >= len(LogFilePathList):
        print("VRChatのログフォルダ " + LogFileDir + " 内にログファイル " + LogFileSearchWord + " を見つけられませんでした")
        input("何かキーを押すとプログラムを終了します")
        sys.exit(0)

    # 名前でソート (ファイルの更新日時でソートするとVRChatを複数起動している時にバグるので諦め)
    LogFilePathList.sort(reverse=True)
    return LogFilePathList[0]


# ファイルの行数を取得する関数
def GetLastLineNumFromLogFile(InputFilePath):
    # ファイルの行数を取得 https://stackoverflow.com/questions/845058/how-to-get-the-line-count-of-a-large-file-cheaply-in-python/1019572#1019572
    with open(InputFilePath, "rb") as f:
        Line = sum([1 for Temp in f])                                           # ファイルの行数分1が格納されたリストを内包表記で生成しSumすることで行数取得
    #print(Line)
    return Line


# ワールド情報を更新する関数
def UpdateNowWorldInfo(s):
    # 引数sとして渡される文字列例
    """
    <GroupPublic>
    wrld_4432ea9b-729c-46e3-8eaf-846aa0a37fdd:00000~group(grp_00000000-0000-0000-0000-000000000000)~groupAccessType(public)~region(use)

    <Group+>
    wrld_4432ea9b-729c-46e3-8eaf-846aa0a37fdd:00000~group(grp_00000000-0000-0000-0000-000000000000)~groupAccessType(plus)~region(us)

    <Group>
    wrld_4432ea9b-729c-46e3-8eaf-846aa0a37fdd:00000~group(grp_00000000-0000-0000-0000-000000000000)~groupAccessType(members)~region(eu)

    <Public>
    wrld_4432ea9b-729c-46e3-8eaf-846aa0a37fdd:00000~region(eu)

    <Freiends+>
    wrld_4432ea9b-729c-46e3-8eaf-846aa0a37fdd:00000~hidden(usr_00000000-0000-0000-0000-000000000000)~region(jp)

    <Freiends>
    wrld_4432ea9b-729c-46e3-8eaf-846aa0a37fdd:00000~friends(usr_00000000-0000-0000-0000-000000000000)~region(jp)

    <Invite+>
    wrld_4432ea9b-729c-46e3-8eaf-846aa0a37fdd:00000~private(usr_00000000-0000-0000-0000-000000000000)~canRequestInvite~region(jp)

    <Invite>
    wrld_4432ea9b-729c-46e3-8eaf-846aa0a37fdd:00000~private(usr_00000000-0000-0000-0000-000000000000)~region(jp)
    """

    global NowWorldInfo                                                         # 関数内で変数の値を変更したい場合はglobalにする

    SplittedWorldIdLog = re.split('[~:]', s)                                    # ~と:で文字列を分割してリストにする

    if 2 <= len(SplittedWorldIdLog):
        # BluePrintIdとInstanceIdを取得
        NowWorldInfo["BluePrintId"] = SplittedWorldIdLog[0]                     # BluePrintId
        NowWorldInfo["InstanceId"]  = SplittedWorldIdLog[1]                     # InstanceId

        # InstanceType取得
        if "hidden" in s:
            NowWorldInfo["InstanceType"] = "Friends+"

        elif "friends" in s:
            NowWorldInfo["InstanceType"] = "Friends"

        elif "private" in s and "canRequestInvite" in s:
            NowWorldInfo["InstanceType"] = "Invite+"

        elif "private" in s:
            NowWorldInfo["InstanceType"] = "Invite"
        
        elif "group" in s and "groupAccessType(public)" in s:
            NowWorldInfo["InstanceType"] = "GroupPublic"

        elif "group" in s and "groupAccessType(plus)" in s:
            NowWorldInfo["InstanceType"] = "Group+"

        elif "group" in s and "groupAccessType(members)" in s:
            NowWorldInfo["InstanceType"] = "Group"
                
        else:
            NowWorldInfo["InstanceType"] = "Public"

        # 文字列中の括弧内の文字列だけ取り出すfindall用引数文字列 https://takake-blog.com/python-regular-expression/
        GetStrInBracketsArgs = "(?<=\().+?(?=\))"
        
        # InstanceOwnerId取得
        NowWorldInfo["InstanceOwnerId"] = "Unknown"
        for i, TempStr in enumerate(SplittedWorldIdLog):
            if "(usr_" in TempStr or "(grp_" in TempStr:
                # 括弧内の文字列だけ取り出す
                NowWorldInfo["InstanceOwnerId"] = re.findall(GetStrInBracketsArgs, SplittedWorldIdLog[i])[0]

        # Region取得
        NowWorldInfo["Region"] = "Unknown"
        for i, TempStr in enumerate(SplittedWorldIdLog):
            if "region(" in TempStr:
                # 括弧内の文字列だけ取り出す
                NowWorldInfo["Region"] = re.findall(GetStrInBracketsArgs, SplittedWorldIdLog[i])[0]
        
        # 別関数でRejoin用のURLを生成する
        GenerateRejoinUrl()
    
    # 何か情報が足りなかったら全てUnknownにする。稀にログにRegionが記載されないときがあった
    else:
        NowWorldInfo["Name"]            = "Unknown"
        NowWorldInfo["BluePrintId"]     = "Unknown"
        NowWorldInfo["InstanceId"]      = "Unknown"
        NowWorldInfo["InstanceType"]    = "Unknown"
        NowWorldInfo["InstanceOwnerId"] = "Unknown"
        NowWorldInfo["Region"]          = "Unknown"
        NowWorldInfo["RejoinUrl"]       = "Unknown"


# 現在のワールド情報Dict内の情報からRejoin用のURLを生成
def GenerateRejoinUrl():
    global NowWorldInfo                                                         # 関数内で変数の値を変更したい場合はglobalにする

    NowWorldInfo["RejoinUrl"] = "Unknown"                                       # Rejoin用URL

    # 現在のワールド情報がFriends+, Friends, Invite+, Invite, Group+, Group の場合
    if ("Unknown" != NowWorldInfo["BluePrintId"] and
        "Unknown" != NowWorldInfo["InstanceId"] and
        "Unknown" != NowWorldInfo["InstanceType"] and
        "Unknown" != NowWorldInfo["InstanceOwnerId"] and
        "Unknown" != NowWorldInfo["Region"]):
        
        # 共通部分生成 ----------------------------------------------------------
        NowWorldInfo["RejoinUrl"]  = "https://vrchat.com/home/launch?worldId=" + NowWorldInfo["BluePrintId"]
        NowWorldInfo["RejoinUrl"] += "&instanceId=" + NowWorldInfo["InstanceId"]

        # InstanceType部分1/2 --------------------------------------------------
        if   "Friends+" == NowWorldInfo["InstanceType"]:
            NowWorldInfo["RejoinUrl"] += "~hidden("
        
        elif "Friends" == NowWorldInfo["InstanceType"]:
            NowWorldInfo["RejoinUrl"] += "~friends("
        
        elif "Invite+" == NowWorldInfo["InstanceType"] or "Invite" == NowWorldInfo["InstanceType"]:
            NowWorldInfo["RejoinUrl"] += "~private("
        
        elif "GroupPublic" == NowWorldInfo["InstanceType"] or "Group+" == NowWorldInfo["InstanceType"] or "Group" == NowWorldInfo["InstanceType"]:
            NowWorldInfo["RejoinUrl"] += "~group("
        
        # InstanceOwnerId部分 --------------------------------------------------
        NowWorldInfo["RejoinUrl"] += NowWorldInfo["InstanceOwnerId"] + ")"

        # InstanceType部分2/2 --------------------------------------------------
        if   "Invite+" == NowWorldInfo["InstanceType"]:
            NowWorldInfo["RejoinUrl"] += "~canRequestInvite"
        
        elif "GroupPublic" == NowWorldInfo["InstanceType"]:
            NowWorldInfo["RejoinUrl"] += "~groupAccessType(public)"
        
        elif "Group+" == NowWorldInfo["InstanceType"]:
            NowWorldInfo["RejoinUrl"] += "~groupAccessType(plus)"
        
        elif "Group" == NowWorldInfo["InstanceType"]:
            NowWorldInfo["RejoinUrl"] += "~groupAccessType(members)"

        # Rejoin部分 -----------------------------------------------------------
        NowWorldInfo["RejoinUrl"] += "~region(" + NowWorldInfo["Region"] + ")"
    
    # 現在のワールド情報が Public の場合
    if ("Unknown" != NowWorldInfo["BluePrintId"] and
        "Unknown" != NowWorldInfo["InstanceId"] and
        "Public"  == NowWorldInfo["InstanceType"] and
        "Unknown" != NowWorldInfo["Region"]):

        NowWorldInfo["RejoinUrl"]  = "https://vrchat.com/home/launch?worldId=" + NowWorldInfo["BluePrintId"]
        NowWorldInfo["RejoinUrl"] += "&instanceId=" + NowWorldInfo["InstanceId"]
        NowWorldInfo["RejoinUrl"] += "~region(" + NowWorldInfo["Region"] + ")"


# ログを保存する
def SaveMyLogFile(NowLogFilePath, NowOutputLog):
    NowLogFileName = os.path.splitext(os.path.basename(NowLogFilePath))[0]          # 現在監視しているログの拡張子なしのファイル名を取得
    MyLogFileSavePath = MyLogFileSaveDir + "/Organized_" + NowLogFileName + ".txt"  # 保存するログファイルパスを生成

    # ログファイルに追記モードで書き込み
    with open(MyLogFileSavePath, 'a', encoding="utf-8") as NowLogFile:
        NowLogFile.write(NowOutputLog)


# 設定ファイルから設定値を取得更新
def UpdateSettingDict():
    global SettingDict                                                          # 関数内で変数の値を変更したい場合はglobalにする

    # 設定ファイルがなかったらエラー出して終了
    if False == os.path.exists(SettingDictFileName):
        print(".exeと同じディレクトリに " + SettingDictFileName + " がありません")
        input("何かキーを押すとプログラムを終了します")
        sys.exit(0)

    # 設定ファイルを開く
    with open(SettingDictFileName, 'r', encoding="utf-8") as SettingFile:
        SettingFilelinesList = SettingFile.read().splitlines()                  # 1行ずつリスト化 (改行コードなし)

        # 設定値Dict
        SettingDict = {"JSettingValue" : 0,                                     # Join通知設定 (無し:0 個別+全員:1 個別のみ:2 全員のみ:3)
                       "LSettingValue" : 0,                                     # Leave通知設定 (無し:0 個別+全員:1 個別のみ:2 全員のみ:3)
                       "JNameList1" : [],                                       # 個別Join通知ユーザー名リスト1
                       "JNameList2" : [],                                       # 個別Join通知ユーザー名リスト2
                       "JNameList3" : [],                                       # 個別Join通知ユーザー名リスト3
                       "JTimeoutNormal" : 0.0,                                  # Join通知表示秒数 全員
                       "JTimeout1" : 0.0,                                       # Join通知表示秒数 リスト1
                       "JTimeout2" : 0.0,                                       # Join通知表示秒数 リスト2
                       "JTimeout3" : 0.0,                                       # Join通知表示秒数 リスト3
                       "LTimeoutNormal" : 0.0,                                  # Leave通知表示秒数 全員
                       "LTimeout1" : 0.0,                                       # Leave通知表示秒数 リスト1
                       "LTimeout2" : 0.0,                                       # Leave通知表示秒数 リスト2
                       "LTimeout3" : 0.0,                                       # Leave通知表示秒数 リスト3
                       "JVolumeNormal" : 0.0,                                   # Join通知音量 全員
                       "JVolume1" : 0.0,                                        # Join通知音量 リスト1
                       "JVolume2" : 0.0,                                        # Join通知音量 リスト2
                       "JVolume3" : 0.0,                                        # Join通知音量 リスト3
                       "LVolumeNormal" : 0.0,                                   # Leave通知音量 全員
                       "LVolume1" : 0.0,                                        # Leave通知音量 リスト1
                       "LVolume2" : 0.0,                                        # Leave通知音量 リスト2
                       "LVolume3" : 0.0,                                        # Leave通知音量 リスト3
                       "LNameList1" : [],                                       # 個別Leave通知ユーザー名リスト1
                       "LNameList2" : [],                                       # 個別Leave通知ユーザー名リスト2
                       "LNameList3" : [],                                       # 個別Leave通知ユーザー名リスト3
                       "PictureRenameSettingValue" : 0                          # 写真のファイル名へのワールド情報追記設定 (無し:0 ワールド名+ワールドID:1 ワールド名のみ:2 ワールドIDのみ:3) ※マルチレイヤ写真はログが出ないので非対応
                       }
        
        IntSettingDictKeyList = []                                              # int型の設定値Dictのキーをまとめたリスト
        FloatSettingDictKeyList = []                                            # float型の設定値Dictのキーをまとめたリスト

        # 設定値DictのValueの初期値からint型かfloat型か判別してキーをリストに分ける
        for Key, Value in SettingDict.items():
            if int == type(Value):
                IntSettingDictKeyList.append(Key)
            if float == type(Value):
                FloatSettingDictKeyList.append(Key)
        #print(IntSettingDictKeyList)
        #print(FloatSettingDictKeyList)

        # 設定ファイルを1行ずつリスト化したリストの要素分ループ
        for LineNum, SettingFileline in enumerate(SettingFilelinesList):
            # ------------------------------------------------------------------
            if "#Join通知設定" in SettingFileline:
                SettingDict["JSettingValue"] = SettingFilelinesList[LineNum + 1]

            if "#Leave通知設定" in SettingFileline:
                SettingDict["LSettingValue"] = SettingFilelinesList[LineNum + 1]
            

            # ------------------------------------------------------------------
            if "#Join通知表示秒数 全員" in SettingFileline:
                SettingDict["JTimeoutNormal"] = SettingFilelinesList[LineNum + 1]
            
            if "#Join通知表示秒数 リスト1" in SettingFileline:
                SettingDict["JTimeout1"] = SettingFilelinesList[LineNum + 1]
            
            if "#Join通知表示秒数 リスト2" in SettingFileline:
                SettingDict["JTimeout2"] = SettingFilelinesList[LineNum + 1]
            
            if "#Join通知表示秒数 リスト3" in SettingFileline:
                SettingDict["JTimeout3"] = SettingFilelinesList[LineNum + 1]
            
            # ------------------------------------------------------------------
            if "#Leave通知表示秒数 全員" in SettingFileline:
                SettingDict["LTimeoutNormal"] = SettingFilelinesList[LineNum + 1]
            
            if "#Leave通知表示秒数 リスト1" in SettingFileline:
                SettingDict["LTimeout1"] = SettingFilelinesList[LineNum + 1]
            
            if "#Leave通知表示秒数 リスト2" in SettingFileline:
                SettingDict["LTimeout2"] = SettingFilelinesList[LineNum + 1]
            
            if "#Leave通知表示秒数 リスト3" in SettingFileline:
                SettingDict["LTimeout3"] = SettingFilelinesList[LineNum + 1]
            
            # ------------------------------------------------------------------
            if "#Join通知音量 全員" in SettingFileline:
                SettingDict["JVolumeNormal"] = SettingFilelinesList[LineNum + 1]
            
            if "#Join通知音量 リスト1" in SettingFileline:
                SettingDict["JVolume1"] = SettingFilelinesList[LineNum + 1]
            
            if "#Join通知音量 リスト2" in SettingFileline:
                SettingDict["JVolume2"] = SettingFilelinesList[LineNum + 1]
            
            if "#Join通知音量 リスト3" in SettingFileline:
                SettingDict["JVolume3"] = SettingFilelinesList[LineNum + 1]
            
            # ------------------------------------------------------------------
            if "#Leave通知音量 全員" in SettingFileline:
                SettingDict["LVolumeNormal"] = SettingFilelinesList[LineNum + 1]
            
            if "#Leave通知音量 リスト1" in SettingFileline:
                SettingDict["LVolume1"] = SettingFilelinesList[LineNum + 1]
            
            if "#Leave通知音量 リスト2" in SettingFileline:
                SettingDict["LVolume2"] = SettingFilelinesList[LineNum + 1]
            
            if "#Leave通知音量 リスト3" in SettingFileline:
                SettingDict["LVolume3"] = SettingFilelinesList[LineNum + 1]


            # ------------------------------------------------------------------
            if "#個別Join通知ユーザー名リスト1" in SettingFileline:
                SettingDict["JNameList1_Index_S"] = LineNum + 1
            
            if "#個別Join通知ユーザー名リスト2" in SettingFileline:
                SettingDict["JNameList1_Index_E"] = LineNum
                SettingDict["JNameList2_Index_S"] = LineNum + 1
            
            if "#個別Join通知ユーザー名リスト3" in SettingFileline:
                SettingDict["JNameList2_Index_E"] = LineNum
                SettingDict["JNameList3_Index_S"] = LineNum + 1
            
            if "#個別Leave通知ユーザー名リスト1" in SettingFileline:
                SettingDict["JNameList3_Index_E"] = LineNum - 1
                SettingDict["LNameList1_Index_S"] = LineNum + 1
            
            if "#個別Leave通知ユーザー名リスト2" in SettingFileline:
                SettingDict["LNameList1_Index_E"] = LineNum
                SettingDict["LNameList2_Index_S"] = LineNum + 1
            
            if "#個別Leave通知ユーザー名リスト3" in SettingFileline:
                SettingDict["LNameList2_Index_E"] = LineNum
                SettingDict["LNameList3_Index_S"] = LineNum + 1
            
            if "#写真のファイル名へのワールド情報追記設定" in SettingFileline:
                SettingDict["LNameList3_Index_E"] = LineNum - 1
                SettingDict["PictureRenameSettingValue"] = SettingFilelinesList[LineNum + 1]

            #print(LineNum, SettingFileline)
        

        # 設定ファイル内の設定値を、文字列から目的の型に変換。できなかったらエラー出して終了
        # int
        for IntSettingDictKey in IntSettingDictKeyList:
            try:
                SettingDict[IntSettingDictKey] = int(SettingDict[IntSettingDictKey])
            except:
                print(SettingDictFileName + " 内で " + IntSettingDictKey + " の設定値として「" + SettingDict[IntSettingDictKey] + "」を整数に変換できませんでした")
                input("何かキーを押すとプログラムを終了します")
                sys.exit(0)
        # float
        for FloatSettingDictKey in FloatSettingDictKeyList:
            try:
                SettingDict[FloatSettingDictKey] = float(SettingDict[FloatSettingDictKey])
            except:
                print(SettingDictFileName + " 内で " + FloatSettingDictKey + " の設定値として「" + SettingDict[FloatSettingDictKey] + "」を小数に変換できませんでした")
                input("何かキーを押すとプログラムを終了します")
                sys.exit(0)
                

        # 内包表記で空の要素を駆逐する https://qiita.com/github-nakasho/items/a08e21e80cbc9761db2f#%E5%86%85%E5%8C%85%E8%A1%A8%E8%A8%98%E3%81%A7%E7%A9%BA%E3%81%AE%E8%A6%81%E7%B4%A0%E3%82%92%E9%A7%86%E9%80%90%E3%81%99%E3%82%8B
        SettingDict["JNameList1"] = [a for a in SettingFilelinesList[SettingDict["JNameList1_Index_S"] : SettingDict["JNameList1_Index_E"]] if a != '']
        SettingDict["JNameList2"] = [a for a in SettingFilelinesList[SettingDict["JNameList2_Index_S"] : SettingDict["JNameList2_Index_E"]] if a != '']
        SettingDict["JNameList3"] = [a for a in SettingFilelinesList[SettingDict["JNameList3_Index_S"] : SettingDict["JNameList3_Index_E"]] if a != '']
        SettingDict["LNameList1"] = [a for a in SettingFilelinesList[SettingDict["LNameList1_Index_S"] : SettingDict["LNameList1_Index_E"]] if a != '']
        SettingDict["LNameList2"] = [a for a in SettingFilelinesList[SettingDict["LNameList2_Index_S"] : SettingDict["LNameList2_Index_E"]] if a != '']
        SettingDict["LNameList3"] = [a for a in SettingFilelinesList[SettingDict["LNameList3_Index_S"] : SettingDict["LNameList3_Index_E"]] if a != '']
        #print(SettingDict)
    

# Joinしたユーザ通知処理実行
def NotifyJoinUserName(JoinUserNameList):
    global Message                                                              # 関数内で変数の値を変更したい場合はglobalにする

    NowDateTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")         # 現在の日付時刻を取得しておく

    # Joinしたユーザー名リストと、設定ファイルに書いておいた個別Join通知ユーザー名リストで、一致したユーザー名をSetとして取得
    CustomJNameHitSet1 = set(SettingDict["JNameList1"]) & set(JoinUserNameList)
    CustomJNameHitSet2 = set(SettingDict["JNameList2"]) & set(JoinUserNameList)
    CustomJNameHitSet3 = set(SettingDict["JNameList3"]) & set(JoinUserNameList)

    JoinUserNameListStr   = ", ".join(JoinUserNameList)                         # リストをカンマ区切りの文字列に変換
    CustomJNameHitSet1Str = ", ".join(CustomJNameHitSet1)                       # Setをカンマ区切りの文字列に変換
    CustomJNameHitSet2Str = ", ".join(CustomJNameHitSet2)                       # Setをカンマ区切りの文字列に変換
    CustomJNameHitSet3Str = ", ".join(CustomJNameHitSet3)                       # Setをカンマ区切りの文字列に変換

    # Join通知 個別
    if 1 == SettingDict["JSettingValue"] or 2 == SettingDict["JSettingValue"]:
        if 0 < len(CustomJNameHitSet1):
            Message["title"] = "Join : リスト1 ユーザー"
            Message["content"] = CustomJNameHitSet1Str
            Message["timeout"] = SettingDict["JTimeout1"]
            Message["volume"] = SettingDict["JVolume1"]
            Message["audioPath"] = JoinSoundPath_List1

            # コンソールに黄色文字でXSOverlay通知履歴を残しておく
            print_color(NowDateTime + " リスト1 ユーザーの join を通知 : " + Message["content"], "yellow")
            ViewXSOverlayNotification()                                         # XSOverlayに通知依頼を出す

        if 0 < len(CustomJNameHitSet2):
            Message["title"] = "Join : リスト2 ユーザー"
            Message["content"] = CustomJNameHitSet2Str
            Message["timeout"] = SettingDict["JTimeout2"]
            Message["volume"] = SettingDict["JVolume2"]
            Message["audioPath"] = JoinSoundPath_List2

            # コンソールに黄色文字でXSOverlay通知履歴を残しておく
            print_color(NowDateTime + " リスト2 ユーザーの join を通知 : " + Message["content"], "yellow")
            ViewXSOverlayNotification()                                         # XSOverlayに通知依頼を出す
        
        if 0 < len(CustomJNameHitSet3):
            Message["title"] = "Join : リスト3 ユーザー"
            Message["content"] = CustomJNameHitSet3Str
            Message["timeout"] = SettingDict["JTimeout3"]
            Message["volume"] = SettingDict["JVolume3"]
            Message["audioPath"] = JoinSoundPath_List3

            # コンソールに黄色文字でXSOverlay通知履歴を残しておく
            print_color(NowDateTime + " リスト3 ユーザーの join を通知 : " + Message["content"], "yellow")
            ViewXSOverlayNotification()                                         # XSOverlayに通知依頼を出す
    
    # Join通知 全員
    if 1 == SettingDict["JSettingValue"] or 3 == SettingDict["JSettingValue"]:
        Message["title"] = "Join"
        Message["content"] = JoinUserNameListStr
        Message["timeout"] = SettingDict["JTimeoutNormal"]
        Message["volume"] = SettingDict["JVolumeNormal"]
        Message["audioPath"] = JoinSoundPath_Normal
        ViewXSOverlayNotification()                                             # XSOverlayに通知依頼を出す


# Leaveしたユーザ通知処理実行
def NotifyLeaveUserName(LeaveUserNameList):
    global Message                                                              # 関数内で変数の値を変更したい場合はglobalにする

    NowDateTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")         # 現在の日付時刻を取得しておく

    # Leaveしたユーザー名リストと、設定ファイルに書いておいた個別Leave通知ユーザー名リストで、一致したユーザー名をSetとして取得
    CustomLNameHitSet1 = set(SettingDict["LNameList1"]) & set(LeaveUserNameList)
    CustomLNameHitSet2 = set(SettingDict["LNameList2"]) & set(LeaveUserNameList)
    CustomLNameHitSet3 = set(SettingDict["LNameList3"]) & set(LeaveUserNameList)

    LeaveUserNameListStr  = ", ".join(LeaveUserNameList)                        # リストをカンマ区切りの文字列に変換
    CustomLNameHitSet1Str = ", ".join(CustomLNameHitSet1)                       # Setをカンマ区切りの文字列に変換
    CustomLNameHitSet2Str = ", ".join(CustomLNameHitSet2)                       # Setをカンマ区切りの文字列に変換
    CustomLNameHitSet3Str = ", ".join(CustomLNameHitSet3)                       # Setをカンマ区切りの文字列に変換

    # Leave通知 個別
    if 1 == SettingDict["LSettingValue"] or 2 == SettingDict["LSettingValue"]:
        if 0 < len(CustomLNameHitSet1):
            Message["title"] = "Leave : リスト1 ユーザー"
            Message["content"] = CustomLNameHitSet1Str
            Message["timeout"] = SettingDict["LTimeout1"]
            Message["volume"] = SettingDict["LVolume1"]
            Message["audioPath"] = LeaveSoundPath_List1

            # コンソールに黄色文字でXSOverlay通知履歴を残しておく
            print_color(NowDateTime + " リスト1 ユーザーの leave を通知 : " + Message["content"], "red")
            ViewXSOverlayNotification()                                         # XSOverlayに通知依頼を出す

        if 0 < len(CustomLNameHitSet2):
            Message["title"] = "Leave : リスト2 ユーザー"
            Message["content"] = CustomLNameHitSet2Str
            Message["timeout"] = SettingDict["LTimeout2"]
            Message["volume"] = SettingDict["LVolume2"]
            Message["audioPath"] = LeaveSoundPath_List2

            # コンソールに黄色文字でXSOverlay通知履歴を残しておく
            print_color(NowDateTime + " リスト2 ユーザーの leave を通知 : " + Message["content"], "red")
            ViewXSOverlayNotification()                                         # XSOverlayに通知依頼を出す
        
        if 0 < len(CustomLNameHitSet3):
            Message["title"] = "Leave : リスト3 ユーザー"
            Message["content"] = CustomLNameHitSet3Str
            Message["timeout"] = SettingDict["LTimeout3"]
            Message["volume"] = SettingDict["LVolume3"]
            Message["audioPath"] = LeaveSoundPath_List3

            # コンソールに黄色文字でXSOverlay通知履歴を残しておく
            print_color(NowDateTime + " リスト3 ユーザーの leave を通知 : " + Message["content"], "red")
            ViewXSOverlayNotification()                                         # XSOverlayに通知依頼を出す
    
    # Leave通知 全員
    if 1 == SettingDict["LSettingValue"] or 3 == SettingDict["LSettingValue"]:
        Message["title"] = "Leave"
        Message["content"] = LeaveUserNameListStr
        Message["timeout"] = SettingDict["LTimeoutNormal"]
        Message["volume"] = SettingDict["LVolumeNormal"]
        Message["audioPath"] = LeaveSoundPath_Normal
        ViewXSOverlayNotification()                                             # XSOverlayに通知依頼を出す


# XSOverlayに通知依頼を出す処理
def ViewXSOverlayNotification():
    SendData = json.dumps(Message).encode("utf-8")                              # 通知文字列をJSON形式にエンコード
    MySocket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)                 # ソケット通信用インスタンスを生成
    MySocket.sendto(SendData, ("127.0.0.1", 42069))                             # XSOverlayに通知依頼を送信
    MySocket.close()                                                            # ソケット通信終了
    # https://zenn.dev/eeharumt/scraps/95f49a62dd809a
    # https://gist.github.com/nekochanfood/fc8017d8247b358154062368d854be9c

    time.sleep(0.1)


# 写真のファイル名にワールド情報を追記する処理
def AddWorldInfomationForPictureFileName():
    time.sleep(0.2)

    # 写真のファイル名へのワールド情報追記設定値 (無し:0 ワールド名+ワールドID:1 ワールド名のみ:2 ワールドIDのみ:3)
    RenameSettingValue = SettingDict["PictureRenameSettingValue"]

    # ワールド名内で、ファイル名に使えない文字列をハイフンに置換する http://iori084.blog.fc2.com/blog-entry-16.html
    ModifiedWorldName = re.sub(r'[\\|/|:|?|"|<|>|\|]', '-', NowWorldInfo["Name"])

    # 写真のパスがまとめられたリストの要素数分ループ
    for PictireFilePath in PictireFilePathList:
        FileExt  = os.path.splitext(PictireFilePath)[1]                         # 拡張子
        NewPictireFilePath =  os.path.splitext(PictireFilePath)[0]              # 拡張子なしのフルパス文字列

        # ワールド名追記
        if 1 == RenameSettingValue or 2 == RenameSettingValue:
            NewPictireFilePath += '_' + ModifiedWorldName
        
        # ワールドID追記
        if 1 == RenameSettingValue or 3 == RenameSettingValue:
            NewPictireFilePath += '_' + NowWorldInfo["BluePrintId"]
        
        # 拡張子追加
        NewPictireFilePath += FileExt

        # リネーム実行
        try:
            os.rename(PictireFilePath, NewPictireFilePath)
        except:
            print("写真 " + PictireFilePath + " のファイル名を " + NewPictireFilePath + " に変更しようとして失敗しました")
            pass





# 初期化
# ==============================================================================================================
# コンソールで文字に色を付けてprintする用の色Dict 参考:Pythonのprintいろいろ 文字色 https://qiita.com/yuto16/items/cb2f812b0d966b6ae920
color_dic = {"black":"\033[30m", "red":"\033[31m", "green":"\033[32m", "yellow":"\033[33m", "blue":"\033[34m", "end":"\033[0m"}

# exe化したコンソールでもANSIカラーを使うため、os.system('')でサブシェルを実行しておくらしい
# 参考:コマンドプロンプトで実行するPythonスクリプトにおけるANSIカラーの出力 https://qiita.com/hermite2053/items/937dcce6f0c31c13ddd7
os.system('')

ExeDir = os.path.dirname(sys.argv[0])                                           # exeが置かれているディレクトリ取得
#print("ExeDir", ExeDir)

# 検索ワードリスト
SearchWordList = [""] * 5                                                       # 検索ワードリストの要素数
SearchWordList[0] = "Entering Room: "                                           # Joinしたワールド名 検索用
SearchWordList[1] = "Joining wrld_"                                             # JoinしたワールドID 検索用
SearchWordList[2] = "OnPlayerJoinComplete "                                     # Joinしたユーザ名 検索用
SearchWordList[3] = "Unregistering "                                            # Leaveしたユーザ名 検索用
SearchWordList[4] = "Took screenshot to: "                                      # Camera, ScreenShotで写真撮った 検索用 (マルチレイヤで撮った写真はログに記録されない)

LogFileDir = os.path.expanduser("~/AppData/LocalLow/VRChat/VRChat/")            # VRChatのログファイルが保存されるディレクトリ
LogFileSearchWord = "output_log*.txt"                                           # VRChatのログファイル名 検索文字列

VRChatWindowName = "VRChat"                                                     # VRChat検出用のウィンドウ名定義
IsFirstLoop = True                                                              # 初回ループ判定フラグ

LogFilePath =  ""                                                               # 最新のログファイルのパス
LogFilePath_Old = ""                                                            # 1ループ前に読み込んだログファイルのパス

LastLineNum = 0                                                                 # 読み込む必要がある行数
LastLineNum_Old = 0                                                             # 読み込み済み行数

# 現在のワールド情報が保存されているDict
NowWorldInfo = {"Name"            : "Unknown",
                "BluePrintId"     : "Unknown",
                "InstanceId"      : "Unknown",
                "InstanceType"    : "Unknown",
                "InstanceOwnerId" : "Unknown",
                "Region"          : "Unknown",
                "RejoinUrl"       : "Unknown"
                }

SettingDictFileName = "設定ファイル.txt"                                         # 設定ファイルのファイル名
SettingDict = {}                                                                # 設定ファイルから取得する設定値がまとめられたDict

# XSOverlay共通通知文 https://xiexe.github.io/XSOverlayDocumentation/#/NotificationsAPI?id=xsoverlay-message-object
Message = {
    "messageType" : 1,          # 1 = Notification Popup, 2 = MediaPlayer Information, will be extended later on.
    "index" : 0,                # Only used for Media Player, changes the icon on the wrist.
    "timeout" : 2.0,            # How long the notification will stay on screen for in seconds
    "height" : 100.0,           # Height notification will expand to if it has content other than a title. Default is 175
    "opacity" : 1.0,            # Opacity of the notification, to make it less intrusive. Setting to 0 will set to 1.
    "volume" : 0.5,             # Notification sound volume.
    "audioPath" : "default",    # File path to .ogg audio file. Can be "default", "error", or "warning". Notification will be silent if left empty.
    "title" : "",               # Notification title, supports Rich Text Formatting
    "useBase64Icon" : False,    # Set to true if using Base64 for the icon image
    "icon" : "default",         # Base64 Encoded image, or file path to image. Can also be "default", "error", or "warning"
    "sourceApp" : "TEST_App"    # Somewhere to put your app name for debugging purposes
    }

# XSOverlayの通知音ファイルのパス定義
JoinSoundPath_Normal = os.path.join(ExeDir, "SoundFile", "JoinSound_Normal.ogg")    # Join通知音 全員
JoinSoundPath_List1  = os.path.join(ExeDir, "SoundFile", "JoinSound_List1.ogg")     # Join通知音 リスト1
JoinSoundPath_List2  = os.path.join(ExeDir, "SoundFile", "JoinSound_List2.ogg")     # Join通知音 リスト2
JoinSoundPath_List3  = os.path.join(ExeDir, "SoundFile", "JoinSound_List3.ogg")     # Join通知音 リスト3

LeaveSoundPath_Normal = os.path.join(ExeDir, "SoundFile", "LeaveSound_Normal.ogg")  # Leave通知音 全員
LeaveSoundPath_List1  = os.path.join(ExeDir, "SoundFile", "LeaveSound_List1.ogg")   # Leave通知音 リスト1
LeaveSoundPath_List2  = os.path.join(ExeDir, "SoundFile", "LeaveSound_List2.ogg")   # Leave通知音 リスト2
LeaveSoundPath_List3  = os.path.join(ExeDir, "SoundFile", "LeaveSound_List3.ogg")   # Leave通知音 リスト3


# 多重起動してたら終了
# ==============================================================================================================
MyExeName = os.path.basename(sys.argv[0])                                       # 自分のexe名を取得 (拡張子付き)
#print("MyExeName", MyExeName)

ProcessHitCount = 0                                                             # 自分を同じ名前のexeがプロセス内に何個あるかカウントする用
for MyProcess in psutil.process_iter():                                         # プロセス一覧取得
    try:
        #print(os.path.basename(MyProcess.exe()))
        if MyExeName == os.path.basename(MyProcess.exe()):                      # 自分を同じ名前のexeだったら
            ProcessHitCount = ProcessHitCount + 1                               # カウントアップ
    except:
        pass

# 単一起動時はexeが2つある(なぜかはわからない)。それを超えていたら多重起動しているということなのでここで終了
if 2 < ProcessHitCount:
    print("多重起動のため終了します")
    time.sleep(5)
    sys.exit(0)


# 設定ファイルを読み込む
# ==============================================================================================================
UpdateSettingDict()
#print(SettingDict)

# 整理したログファイルを保存するフォルダを生成。すでにあったらそのまま
# ==============================================================================================================
MyLogFileSaveDir = os.path.join(ExeDir, "ログまとめ")
#print("MyLogFileSaveDir", MyLogFileSaveDir)
os.makedirs(MyLogFileSaveDir, exist_ok=True)                                    # フォルダを生成。すでにあったらそのまま


# ログファイル監視処理
# ==============================================================================================================
print("カスタムJoin通知 v0.0")
print("----------------------------------------------------------------")

while True:
    time.sleep(1)

    PrintLog = ""                                                               # コンソール表示用文字列
    OutputLog = ""                                                              # ログ保存用文字列
    NowJoinUserNameList = []                                                    # joinしたユーザー名のリスト
    NowLeaveUserNameList = []                                                   # Leaveしたユーザー名のリスト
    PictireFilePathList = []                                                    # 撮影した写真のパスリスト

    # VRChatが起動してなかったら、処理をスキップ ----------------------------------
    # VRChat検出用のウィンドウ名と完全一致するウィンドウを取得してみる。なかった場合は0(int)、あった場合はウィンドウハンドル(int)が返ってくる
    WindowHandle = win32gui.FindWindow(None, VRChatWindowName)
    # VRChatが起動してなかった
    if 0 == WindowHandle:
        continue
        
    # 初回ループ時実行 ----------------------------------------------------------
    if True == IsFirstLoop:
        LogFilePath_Old = GetNewestLogFilePath()                                # 現在のログファイルパスを保存しておく
        LastLineNum_Old = GetLastLineNumFromLogFile(LogFilePath_Old)            # 行数を取得
        IsFirstLoop = False                                                     # 初回ループ判定フラグを折る

    # 通常処理 ------------------------------------------------------------------
    LogFilePath = GetNewestLogFilePath()                                        # 新しいログファイルパスを取得
    LastLineNum = GetLastLineNumFromLogFile(LogFilePath)                        # 行数を取得
    #print(LastLineNum)

    # 新しくログファイルが生成されてたら読み込み済み行数をリセットする (VRChatを再度起動した場合想定)
    if LogFilePath_Old != LogFilePath:
        LastLineNum_Old = 0                                                     # 読み込み済み行数リセット

        # ワールド情報Dictリセット
        NowWorldInfo["Name"]            = "Unknown"
        NowWorldInfo["BluePrintId"]     = "Unknown"
        NowWorldInfo["InstanceId"]      = "Unknown"
        NowWorldInfo["InstanceType"]    = "Unknown"
        NowWorldInfo["InstanceOwnerId"] = "Unknown"
        NowWorldInfo["Region"]          = "Unknown"
        NowWorldInfo["RejoinUrl"]       = "Unknown"

    # ログファイルの行数が変わってなかったらここで処理をスキップ
    if LastLineNum_Old == LastLineNum:
        continue

    # ログファイルを開く
    with open(LogFilePath, 'r', encoding="utf-8") as f:
        LineList_Iterator = islice(f, LastLineNum_Old, LastLineNum)             # 読み込む行目は0始まり
        LineList = [str(t).rstrip() for t in LineList_Iterator]                 # イテレータリストからstrリストに変換。文末の改行も消去
    
        # LineListを1行ずつLineに読み込み、SearchWordList内の検索ワード(SearchWord)が存在したら、整理してHitLineListに追加する
        # 0列目に時間部分(19文字目まで)、1列目にSearchWordListの何番目にヒットしたか、2列目に必要なログ部分(46文字目から)を格納
        HitLineList = [[Line[:19], N, Line[34:]] for Line in LineList for N, SearchWord in enumerate(SearchWordList) if SearchWord in Line]
        
        for j in range(len(HitLineList)):
            # HitDataList 0列目の時間をstrからdatetime型に変換
            HitLineList[j][0] = datetime.datetime.strptime(HitLineList[j][0], "%Y.%m.%d %H:%M:%S")           
            #print(*HitLineList[j])

            # Joinしたワールド名を取得できた
            if   0 == HitLineList[j][1]:
                NowWorldInfo["Name"] = HitLineList[j][2][27:]
            
            # JoinしたワールドIDを取得できた
            elif 1 == HitLineList[j][1]:
                UpdateNowWorldInfo(HitLineList[j][2][20:])                      # ワールド情報Dict更新

                # ログ保存用文字列生成
                OutputLog += str(HitLineList[j][0]) + " Entering World\n"
                OutputLog += "----------------------------------------------------------------\n"
                OutputLog += "World Name        : " + NowWorldInfo["Name"]            + "\n"
                OutputLog += "World ID          : " + NowWorldInfo["BluePrintId"]     + "\n"
                OutputLog += "Instance Id       : " + NowWorldInfo["InstanceId"]      + "\n"
                OutputLog += "Instance Type     : " + NowWorldInfo["InstanceType"]    + "\n"
                OutputLog += "Instance Owner Id : " + NowWorldInfo["InstanceOwnerId"] + "\n"
                OutputLog += "Region            : " + NowWorldInfo["Region"]          + "\n"
                OutputLog += "Rejoin URL        : " + NowWorldInfo["RejoinUrl"]       + "\n"
                OutputLog += "----------------------------------------------------------------\n"
                
                # コンソール表示用文字列生成
                PrintLog = copy.copy(OutputLog)
            
            # Joinしたユーザ名を取得できた
            elif 2 == HitLineList[j][1]:
                # ログ保存用文字列 追記
                OutputLog += str(HitLineList[j][0]) + " Join  : " + HitLineList[j][2][33:] + "\n"
                # コンソール表示用文字列 追記
                PrintLog  += str(HitLineList[j][0]) + color_dic["green"] + " Join" + color_dic["end"] + "  : " + HitLineList[j][2][33:] + "\n"
                # Join通知用ユーザ名リストにユーザ名 追記
                NowJoinUserNameList.append(HitLineList[j][2][33:])
            
            # Leaveしたユーザ名を取得できた
            elif 3 == HitLineList[j][1]:
                # ログ保存用文字列 追記
                OutputLog += str(HitLineList[j][0]) + " Leave : " + HitLineList[j][2][26:] + "\n"
                # コンソール表示用文字列 追記
                PrintLog  += str(HitLineList[j][0]) + color_dic["red"] + " Leave" + color_dic["end"] + " : " + HitLineList[j][2][26:] + "\n"
                # Leave通知用ユーザ名リストにユーザ名 追記
                NowLeaveUserNameList.append(HitLineList[j][2][26:])
            
            # 撮影した写真のパスを取得できた
            elif 4 == HitLineList[j][1]:                     
                PictireFilePathList.append(HitLineList[j][2][33:])              # リネーム用に撮影した写真のパスをリストに追加
                PictireFileName = os.path.basename(HitLineList[j][2][33:])      # 撮影した写真のフルパスから拡張子付きのファイル名取得
                # ログ保存用文字列 追記
                OutputLog += str(HitLineList[j][0]) + " Take Picture : " + PictireFileName + "\n"
                # コンソール表示用文字列 追記
                PrintLog  += copy.copy(OutputLog)

            #print(*HitLineList[j])
    
    # 必要があったらログを表示
    if 0 < len(PrintLog):
        print(PrintLog, end="")

    # 必要があったらログを保存
    if 0 < len(OutputLog):
        SaveMyLogFile(LogFilePath, OutputLog)
    
    # Joinしたユーザが居たら通知処理実行
    if 0 < len(NowJoinUserNameList):
        NotifyJoinUserName(NowJoinUserNameList)
    
    # Leaveしたユーザが居たら通知処理実行
    if 0 < len(NowLeaveUserNameList):
        NotifyLeaveUserName(NowLeaveUserNameList)
    
    # 写真のファイル名にワールド情報を追記する処理
    if 0 < len(PictireFilePathList) and 0 != SettingDict["PictureRenameSettingValue"]:
        AddWorldInfomationForPictureFileName()


    LogFilePath_Old = LogFilePath                                               # 現在のログファイルパスを保存しておく
    LastLineNum_Old = LastLineNum                                               # 現在のログファイル読み込み済み行数を保存しておく



