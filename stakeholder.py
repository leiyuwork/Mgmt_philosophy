import pandas as pd  # excel data操作要パッケージを導入
# 顧客、株主、従業員、地域社会　各ステークホルダー群のキーワードを特定
sh_dic_customer = ["国民", "人々", "取引先", "顧客", "お客", "クライアント", "消費者", "カスタマー", "需要家", "得意先", "得意様", "ユーザ", "エンドユーザー",
                   "ファン", "企業様", "利用者", "オーナー", "入居者", "不動産所有者", "患者", "施主", "工事業者", "旅行者"]
sh_dic_share = ["株主", "ストックホルダー", "出資者", "投資家"]
sh_dic_staff = ["従業員", "社員", "メンバー", "スタッフ", "クルー", "働く人"]
sh_dic_community = ["地球", "社会", "地域", "共同体", "コミュニティ", "自治体", "ＳＤＧｓ", "人類"]

IS_data = pd.DataFrame(pd.read_excel(r"C:\Users\Ray94\Downloads\結果＿ステークホルダー言及順.xlsx"))  # 経営理念記載の作業用excelを読み込む（全社分）
Item_value = IS_data.loc[:, ['記載内容']].values.tolist()  # 経営理念記載の列を用意する（全社分）
for Items in Item_value:  # 行/会社ごとに分析する。会社数の分だけ以下の処理を繰り返す
    for Philosophy in Items:  # 経営理念の文字を取り出す（一社分）
        final = {}  # 最終結果の箱を用意
        # まずは顧客のキーワードを言及したか否かを分析する。複数言及した場合、最初のものを「顧客」の結果として残す
        customer = {}    # 顧客グループの結果の箱を用意
        for sh in sh_dic_customer:   # 顧客グループのキーワード一つ一つ、以下の処理を行う
            sh_num = Philosophy.find(sh) + 1   # 経営理念の何文字目にキーワードを言及したかをfind
            customer[sh] = Philosophy.find(sh) + 1   # "お客", "顧客"のように複数言及する場合もあるので、全部箱に入れる
        try:
            # 言及していないキーワードを無視して、複数言及した場合、最初のものを「顧客」の結果として残す
            key_customer = sorted([(k, v) for k, v in customer.items() if v > 0], key=lambda x: x[1])[0]  #言及していないキーワードを無視して
            final["customer"] = key_customer[1]
        except:
            # 一つも言及していない場合は9999999の値を付与する
            final["customer"] = 999999

        # 従業員の分析
        staff = {}
        for sh in sh_dic_staff:
            sh_num = Philosophy.find(sh) + 1
            staff[sh] = Philosophy.find(sh) + 1
        try:
            key_staff = sorted([(k, v) for k, v in staff.items() if v > 0], key=lambda x: x[1])[0]
            final["staff"] = key_staff[1]
        except:
            final["staff"] = 999999

        # 株主の分析
        share = {}
        for sh in sh_dic_share:
            sh_num = Philosophy.find(sh) + 1
            share[sh] = Philosophy.find(sh) + 1
        try:
            key_share = sorted([(k, v) for k, v in share.items() if v > 0], key=lambda x: x[1])[0]
            final["share"] = key_share[1]
        except:
            final["share"] = 999999

        # 地域社会の分析
        community = {}
        for sh in sh_dic_community:
            sh_num = Philosophy.find(sh) + 1
            community[sh] = Philosophy.find(sh) + 1
        try:
            key_community = sorted([(k, v) for k, v in community.items() if v > 0], key=lambda x: x[1])[0]
            final["community"] = key_community[1]
        except:
            final["community"] = 999999

        # 四つのグループの結果を言及した順にソートする、DataFrameを加工して、出力すること
        final = sorted(final.items(), key=lambda x: x[1])
        final = [final]
        print(final)
        Item = pd.DataFrame(final)
        Item.to_csv(r"C:\Users\Ray94\Downloads\111.csv", mode='a', index=False, header=None)
