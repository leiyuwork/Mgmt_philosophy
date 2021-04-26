import pandas as pd  # excel data操作要パッケージを導入

# 顧客、株主、従業員、地域社会　各ステークホルダー群のキーワードを特定
sh_dic_customer = ["エンドユーザ", "オーナー ", "お客 ", "お得意", "カスタマー ", "クライアント ", "小売業者", "ファン ", "ユーザ ",
                   "会員", "患者 ", "企業様 ", "顧客", "工事業者 ", "国民", "市場ニーズ", "施主 ", "飼い主さま", "事業者さま",
                   "取引先 ", "取引企業様", "需要家 ", "消費者 ", "生活者", "賃借人", "得意先 ", "得意様 ", "入居者 ", "発注者",
                   "不動産所有者 ", "利用者 ", "旅行者", "サービス提供先", "顧客", "協業先", "協力先", "提携先", "派遣先", "テナント",
                   "ニーズ", "居住者", "使用者", "視聴者", "潜在的欲求", "お取引業者様", "お取引様"]

sh_dic_provider = ["サービス提供会社", "サービス提供企業", "サービス提供者", "サプライヤー", "サポート会社", "サポート企業", "パートナー", "卸売業者", "家主様", "関係会社",
                   "協業会社", "協力会社", "協力企業", "協力業者", "協力工場", "協力者", "相手先企業", "提供会社", "提供者", "提携メーカー", "提携会社", "提携企業",
                   "提携協力企業", "提携代理店", "発注者", "サプライヤー", "部品メーカー", "供給メーカー", "原材料メーカー", "供給業者", "仕入先"]

sh_dic_staff = ["アルバイト", "エンジニア", "クルー", "スタッフ", "チームワーク", "チーム力", "メンバー", "安全環境", "安全管理", "安全教育", "安全対策", "育成制度",
                "管理職", "企業市民", "給与水準", "教育訓練", "教育研修", "教育制度", "勤務環境", "勤務管理", "勤務体系", "経営陣以下", "採用", "自己実現", "社員",
                "就労環境", "従業員", "従事者", "職員", "職場", "人づくり", "人材", "人財", "組織活性化", "仲間", "働きがい", "働き甲斐", "働き方", "働く人", "同僚",
                "販売員", "福利厚生", "役員", "役職員", "労災ゼロ", "労使一体", "労使相互", "労働安全衛生", "労働環境", "労働関係法令", "労働災害", "労働時間", "労働条件",
                "労働人権", "有給休暇", "健康経営", "雇用制度", "自己開発", "自己規律", "自己啓発", "自己研鑽"]

sh_dic_shareholder = ["株主", "企業価値", "資本市場", "投資者", "利益還元", "ストックホルダー", "出資者", "投資家", "投資主", "還元策", "還元利益", "配当"]

sh_dic_competitor = ["競業他社", "競合", "競争会社", "同業者", "同業他社"]

sh_dic_authority = ["官公庁", "県庁", "府庁", "同庁", "都庁", "市役所", "区役所", "関係者機関", "規制当局", "経済産業省", "厚生労働省", "行政等", "国家機関",
                    "国土交通省", "自治体", "政府", "地方公共団体", "内閣", "日本政府", "当局"]

sh_dic_community = ["コミュニティ", "まちづくり", "街づくり", "共同体", "地域価値", "地域課題", "地域拡大", "地域格差", "地域活性化", "地域完結型", "地域環境", "地域関係者",
                    "地域企業", "地域機関", "地域規模", "地域共栄", "地域共生", "地域経済", "地方創生", "地域顧客", "地域公共インフラ", "地域貢献", "地域催事", "地域最良",
                    "地域産業", "地域社会", "地域商店街", "地域住民", "地域重視", "地域循環型社会", "地域消費者", "地域情報", "地域情報誌", "地域振興", "地域性", "地域生活",
                    "地域生活者", "地域戦略", "地域全体", "地域創生", "地域的", "地域展開", "地域特性", "地域内シェア", "地域発展", "地域販売戦略", "地域文化", "地域別",
                    "地域包括ケア", "地域毎", "地域密着", "地域要因", "地域流通", "地域歴史", "地元ネットワーク", "地元", "地産地消", "地場", "地方経済"]

sh_dic_society = ["共栄社会", "共生トライアングル", "共生環境", "共生共存", "共生社会", "共生創造企業", "共存共栄", "共存共生", "共存協調", "共存同栄", "地球環境", "共通価値",
                  "共通課題", "国際社会", "社会環境", "社会貢献"]

IS_data = pd.DataFrame(pd.read_excel(r"C:\Users\Ray94\Downloads\結果＿ステークホルダー言及順.xlsx"))  # 経営理念記載の作業用excelを読み込む（全社分）
Item_value = IS_data.loc[:, ['記載内容']].values.tolist()  # 経営理念記載の列を用意する（全社分）
for Items in Item_value:  # 行/会社ごとに分析する。会社数の分だけ以下の処理を繰り返す
    for Philosophy in Items:  # 経営理念の文字を取り出す（一社分）
        final = {}  # 最終結果の箱を用意
        # まずは顧客のキーワードを言及したか否かを分析する。複数言及した場合、最初のものを「顧客」の結果として残す
        customer = {}  # 顧客グループの結果の箱を用意
        for sh in sh_dic_customer:  # 顧客グループのキーワード一つ一つ、以下の処理を行う
            sh_num = Philosophy.find(sh) + 1  # 経営理念の何文字目にキーワードを言及したかをfind
            customer[sh] = Philosophy.find(sh) + 1  # "お客", "顧客"のように複数言及する場合もあるので、全部箱に入れる
        try:
            # 言及していないキーワードを無視して、複数言及した場合、最初のものを「顧客」の結果として残す
            key_customer = sorted([(k, v) for k, v in customer.items() if v > 0], key=lambda x: x[1])[
                0]  # 言及していないキーワードを無視して
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
        shareholder = {}
        for sh in sh_dic_shareholder:
            sh_num = Philosophy.find(sh) + 1
            shareholder[sh] = Philosophy.find(sh) + 1
        try:
            key_shareholder = sorted([(k, v) for k, v in shareholder.items() if v > 0], key=lambda x: x[1])[0]
            final["shareholder"] = key_shareholder[1]
        except:
            final["shareholder"] = 999999

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
