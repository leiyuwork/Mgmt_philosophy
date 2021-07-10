import pandas as pd  # excel data操作要パッケージを導入

# 顧客、株主、従業員、地域社会　各ステークホルダー群のキーワードを特定
sh_dic_customer = ["お取引業者様", "お取引様", "エンドユーザ", "人々", "オーナー", "お客", "お得意", "カスタマー", "クライアント", "小売業者", "ファン", "ユーザ", "会員",
                   "患者", "企業様", "顧客", "工事業者", "国民", "市場ニーズ", "施主", "飼い主さま", "事業者さま", "取引先", "取引企業様", "需要家", "消費者",
                   "生活者", "賃借人", "得意先", "得意様", "入居者", "発注者", "不動産所有者", "利用者", "旅行者", "顧客", "協業先", "協力先", "提携先", "派遣先",
                   "テナント", "ニーズ", "居住者", "使用者", "視聴者", "潜在的欲求"]

sh_dic_supplier = ["発注者", "サービス提供企業", "サポート会社", "サポート企業", "パートナー", "卸売業者", "家主様", "協業会社", "協力会社", "協力企業", "協力業者",
                   "協力工場", "協力者", "相手先企業", "提携メーカー", "提携会社", "提携企業", "提携協力企業", "提携代理店", "仕入先"]

sh_dic_staff = ["安全環境", "安全管理", "安全対策", "アルバイト", "エンジニア", "クルー", "スタッフ", "チームワーク", "チーム力", "メンバー", "安全教育", "育成制度",
                "管理職", "企業市民", "給与水準", "教育訓練", "教育研修", "教育制度", "勤務環境", "勤務管理", "勤務体系", "経営陣以下", "採用", "自己実現", "社員",
                "就労環境", "従業員", "従事者", "職員", "職場", "職場づくり", "職場環境", "人づくり", "人材", "人材育成", "人財", "組織活性化", "仲間", "働きがい",
                "働き甲斐", "働き方", "働く人", "同僚", "販売員", "福利厚生", "役員", "役職員", "労災ゼロ", "労使一体", "労使相互", "労働安全衛生", "労働環境",
                "労働関係法令", "労働災害", "労働時間", "労働条件", "労働人権", "有給休暇", "健康経営", "雇用制度", "自己開発", "自己規律", "自己啓発", "自己研鑽"]

sh_dic_shareholder = ["資本市場", "株主", "企業価値", "投資者", "利益還元", "ストックホルダー", "出資者", "投資家", "投資主", "還元策", "還元利益", "配当"]

sh_dic_competitor = ["競業他社", "同業他社", "他社", "競合", "競争会社", "同業者"]

sh_dic_government = ["官公庁", "県庁", "府庁", "同庁", "都庁", "市役所", "区役所", "関係者機関", "規制当局", "経済産業省", "厚生労働省", "行政等", "国家機関",
                     "国土交通省", "自治体", "政府", "地方公共団体", "内閣", "日本政府", "当局"]

sh_dic_community = ["コミュニティ", "まちづくり", "街づくり", "共同体", "地域", "地域価値", "地域課題", "地域拡大", "地域格差", "地域活性化", "地域完結型", "地域環境",
                    "地域関係者", "地域企業", "地域機関", "地域規模", "地域共栄", "地域共生", "地域経済", "地方創生", "地域顧客", "地域公共インフラ", "地域貢献", "地域催事",
                    "地域最良", "地域産業", "地域社会", "地域商店街", "地域住民", "地域重視", "地域循環型社会", "地域消費者", "地域情報", "地域情報誌", "地域振興", "地域性",
                    "地域生活", "地域生活者", "地域戦略", "地域全体", "地域創生", "地域的", "地域展開", "地域特性", "地域内シェア", "地域発展", "地域販売戦略", "地域文化",
                    "地域別", "地域包括ケア", "地域毎", "地域密着", "地域要因", "地域流通", "地域歴史", "地元ネットワーク", "地元", "地産地消", "地場", "地方経済"]

sh_dic_society = ["共栄社会", "共生トライアングル", "共生環境", "共生共存", "共生社会", "共生創造企業", "共存共栄", "共存共生", "共存協調", "共存同栄", "地球環境", "共通価値",
                  "共通課題", "国際社会", "社会環境", "社会貢献"]

sh_dic_international = ["ＡＳＥＡＮ", "アジア", "アセアン", "アゼルバイジャン", "アフガニスタン", "インドネシア", "ウズベキスタン", "カザフスタン", "カンボジア", "シンガポール",
                        "パキスタン", "バングラデシュ", "ベトナム", "マニラ", "ミャンマー", "ラオス", "韓国", "香港", "台湾", "東ティモール", "北朝鮮", "インド",
                        "シンガポール", "スリランカ", "フィリピン", "マレーシア", "キルギス", "タイ", "ネパール", "ブータン", "マカオ", "タジキスタン", "トルクメニスタン",
                        "ブルネイ", "モルディブ", "モンゴル", "中国", "アフリカ", "アルジェリア", "ウガンダ", "エチオピア", "ガーナ", "ガイアナ", "ガボン", "カメルーン",
                        "ガンビア", "ギニア", "赤道ギニア", "中央アフリカ共和国", "南アフリカ", "南スーダン", "コートジボワール", "コンゴ", "コンゴ共和国", "ザンジバル",
                        "タンザニア", "ジンバブエ", "スーダン", "ソマリア", "タンザニア", "チュニジア", "ナイジェリア", "モロッコ", "オーストラリア", "オセアニア",
                        "ニューカレドニア", "ニュージーランド", "パプアニューギニア", "パラオ", "フィジー", "豪州", "ソロモン諸島", "ミクロネシア連邦", "キリバス",
                        "フランス領ポリネシア", "アメリカ", "日米", "米国", "北米", "カナダ", "欧米", "欧米", "アイスランド", "アイルランド", "アルゼンチン",
                        "アルバニア", "アルメニア", "アンゴラ", "アンドラ", "英国", "イギリス", "イタリア", "ウクライナ", "エストニア", "オーストリア", "オランダ",
                        "キプロス", "ギリシャ", "ノルウェー", "ハンガリー", "フランス", "ベルギー", "ポーランド", "ボスニア・ヘルツェゴビナ", "ポルトガル", "モナコ",
                        "ヨーロッパ", "ラトビア", "リトアニア", "リヒテンシュタイン", "ルーマニア", "ルクセンブルク", "欧州", "北マケドニア", "グリーンランド", "サンマリノ",
                        "スイス", "スウェーデン", "スペイン", "スロバキア", "スロベニア", "セルビア", "チェコ", "デンマーク", "ドイツ", "トルコ", "フィンランド",
                        "グルジア", "クロアチア", "ブルガリア", "マルタ", "モルドバ", "モンテネグロ", "コソボ", "バミューダ諸島", "ベラルーシ", "サモア", "トンガ",
                        "ナウル", "バヌアツ", "マーシャル諸島", "ロシア", "ケニア", "コモロ", "サントメ・プリンシペ", "ザンビア", "シエラレオネ", "ジブチ", "セーシェル",
                        "セネガル", "チャド", "トーゴ", "ナミビア", "ニカラグア", "ニジェール", "ブルキナファソ", "ブルンジ", "ベナン", "ボツワナ", "マダガスカル",
                        "マラウイ", "マリ", "モーリシャス", "モーリタニア", "モザンビーク", "リベリア", "ルワンダ", "レソト", "アラブ首長国連邦", "イエメン", "イスラエル",
                        "イラク", "イラン", "エジプト", "オマーン", "カタール", "パレスチナ自治区", "ヨルダン", "リビア", "クウェート", "サウジアラビア", "シリア",
                        "バーレーン", "レバノン", "メキシコ", "ウルグアイ", "エクアドル", "エルサルバドル", "キューバ", "キュラソー", "グアテマラ", "ベネズエラ", "南米",
                        "ジャマイカ", "チリ", "プエルトリコ", "ブラジル", "ペルー", "ホンジュラス", "グレナダ", "ケイマン諸島", "コスタリカ", "コロンビア",
                        "シント・マールテン", "スリナム", "セントルシア", "ドミニカ", "トリニダード・トバゴ", "ハイチ", "パナマ", "バハマ", "パラグアイ", "バルバドス",
                        "ベリーズ", "ボリビア", "日米欧三極", "カ国", "ヵ国", "ヶ国", "各国", "海外", "外国", "グローバル", "クロスボーダー", "環日本海経済圏",
                        "太平洋地域", "新興国", "先進国", "国外", "国内外", "世界", "人類", "世界トップ", "世界ナンバーワン", "世界ブランド", "世界各国", "世界各地",
                        "世界企業", "世界規模", "世界経済", "世界光学工業界", "世界貢献", "世界最悪最大", "世界最強", "世界最高", "世界最先端", "世界市場", "世界視野",
                        "世界社会", "世界人類", "世界水準", "世界数十か国", "世界戦略", "世界全域", "世界全市場", "世界全体", "世界第一級", "世界中", "世界的",
                        "世界標準", "全世界", "国際"]

sh_dic_japan = ["熊本", "群馬", "広島", "香川", "高知", "愛知", "愛媛", "茨城", "岡山", "沖縄", "岩手", "岐阜", "宮崎", "宮城", "京都", "佐賀", "埼玉",
                "三重", "山形", "山口", "山梨", "滋賀", "鹿児島", "秋田", "新潟", "神奈川", "青森", "静岡", "石川", "千葉", "大分", "大阪", "長崎", "長野",
                "鳥取", "島根", "東京", "徳島", "栃木", "奈良", "富山", "福井", "福岡", "福島", "兵庫", "北海道", "和歌山", "関西", "関東", "都区部",
                "都市部", "都心", "信越地区", "日本橋", "相模原市", "堺市", "宇部市", "横浜市", "北九州市", "九州", "札幌", "八戸", "新宿", "政令指定都市", "中四国",
                "四大都市圏", "中部", "東海", "近畿", "京阪神", "東北", "府下", "中京圏", "名古屋", "首都圏", "甲信越", "上越", "北陸", "北関東", "沿線",
                "わが国", "日米", "日米欧三極", "日本", "国内", "県市町村", "国内外", "全国"]

IS_data = pd.DataFrame(pd.read_excel(r"C:\Users\Ray94\Downloads\結果＿ステークホルダー言及順.xlsx"))  # 経営理念記載の作業用excelを読み込む（全社分）
Item_value = IS_data.loc[:, ['記載内容']].values.tolist()  # 経営理念記載の列を用意する（全社分）
for Items in Item_value:  # 行/会社ごとに分析する。会社数の分だけ以下の処理を繰り返す
    for Philosophy in Items:  # 経営理念の文字を取り出す（一社分）
        final = {}  # 最終結果の箱を用意
        final["len"] = len(Philosophy)
        # まずは主体_顧客（下流取引先）のキーワードを言及したか否かを分析する。複数言及した場合、最初のものを「顧客」の結果として残す
        customer = {}  # 主体_顧客（下流取引先）グループの結果の箱を用意
        for sh in sh_dic_customer:  # 主体_顧客（下流取引先）グループのキーワード一つ一つ、以下の処理を行う
            sh_num = Philosophy.find(sh) + 1  # 経営理念の何文字目にキーワードを言及したかをfind
            customer[sh] = Philosophy.find(sh) + 1  # "お客", "顧客"のように複数言及する場合もあるので、全部箱に入れる
        try:
            key_customer = sorted([(k, v) for k, v in customer.items() if v > 0], key=lambda x: x[1])[
                0]  # 言及していないキーワードを無視して
            final["customer"] = key_customer[1]  # 複数言及した場合、言及箇所が最初のものを「主体_顧客（下流取引先）」の結果として残す
            # 複数言及した場合、数を数える
            final["customer_len"] = len(sorted([(k, v) for k, v in customer.items() if v > 0], key=lambda x: x[1]))
        except:
            # 一つも言及していない場合は9999999の値を付与する、言及回数は0となる
            final["customer"] = 999999
            final["customer_len"] = 0

        # 主体_供給業者（上流取引先）の分析
        supplier = {}
        for sh in sh_dic_supplier:
            sh_num = Philosophy.find(sh) + 1
            supplier[sh] = Philosophy.find(sh) + 1
        try:
            key_supplier = sorted([(k, v) for k, v in supplier.items() if v > 0], key=lambda x: x[1])[0]
            final["supplier"] = key_supplier[1]
            final["supplier_len"] = len(sorted([(k, v) for k, v in supplier.items() if v > 0], key=lambda x: x[1]))
        except:
            final["supplier"] = 999999
            final["supplier_len"] = 0

        # 従業員の分析
        staff = {}
        for sh in sh_dic_staff:
            sh_num = Philosophy.find(sh) + 1
            staff[sh] = Philosophy.find(sh) + 1
        try:
            key_staff = sorted([(k, v) for k, v in staff.items() if v > 0], key=lambda x: x[1])[0]
            final["staff"] = key_staff[1]
            final["staff_len"] = len(sorted([(k, v) for k, v in staff.items() if v > 0], key=lambda x: x[1]))
        except:
            final["staff"] = 999999
            final["staff_len"] = 0

        # 競合の分析
        competitor = {}
        for sh in sh_dic_competitor:
            sh_num = Philosophy.find(sh) + 1
            competitor[sh] = Philosophy.find(sh) + 1
        try:
            key_competitor = sorted([(k, v) for k, v in competitor.items() if v > 0], key=lambda x: x[1])[0]
            final["competitor"] = key_competitor[1]
            final["competitor_len"] = len(sorted([(k, v) for k, v in competitor.items() if v > 0], key=lambda x: x[1]))
        except:
            final["competitor"] = 999999
            final["competitor_len"] = 0

        # 株主の分析
        shareholder = {}
        for sh in sh_dic_shareholder:
            sh_num = Philosophy.find(sh) + 1
            shareholder[sh] = Philosophy.find(sh) + 1
        try:
            key_shareholder = sorted([(k, v) for k, v in shareholder.items() if v > 0], key=lambda x: x[1])[0]
            final["shareholder"] = key_shareholder[1]
            final["shareholder_len"] = len(
                sorted([(k, v) for k, v in shareholder.items() if v > 0], key=lambda x: x[1]))
        except:
            final["shareholder"] = 999999
            final["shareholder_len"] = 0

        # 自治体・政府の分析
        government = {}
        for sh in sh_dic_government:
            sh_num = Philosophy.find(sh) + 1
            government[sh] = Philosophy.find(sh) + 1
        try:
            key_government = sorted([(k, v) for k, v in government.items() if v > 0], key=lambda x: x[1])[0]
            final["government"] = key_government[1]
            final["government_len"] = len(sorted([(k, v) for k, v in government.items() if v > 0], key=lambda x: x[1]))
        except:
            final["government"] = 999999
            final["government_len"] = 0

        # 地域社会の分析
        community = {}
        for sh in sh_dic_community:
            sh_num = Philosophy.find(sh) + 1
            community[sh] = Philosophy.find(sh) + 1
        try:
            key_community = sorted([(k, v) for k, v in community.items() if v > 0], key=lambda x: x[1])[0]
            final["community"] = key_community[1]
            final["community_len"] = len(sorted([(k, v) for k, v in community.items() if v > 0], key=lambda x: x[1]))
        except:
            final["community"] = 999999
            final["community_len"] = 0

        # 社会の分析
        society = {}
        for sh in sh_dic_society:
            sh_num = Philosophy.find(sh) + 1
            society[sh] = Philosophy.find(sh) + 1
        try:
            key_society = sorted([(k, v) for k, v in society.items() if v > 0], key=lambda x: x[1])[0]
            final["society"] = key_society[1]
            final["society_len"] = len(sorted([(k, v) for k, v in society.items() if v > 0], key=lambda x: x[1]))
        except:
            final["society"] = 999999
            final["society_len"] = 0

        #海外の分析
        international = {}
        for sh in sh_dic_international:
            sh_num = Philosophy.find(sh) + 1
            international[sh] = Philosophy.find(sh) + 1
        try:
            key_international = sorted([(k, v) for k, v in international.items() if v > 0], key=lambda x: x[1])[0]
            final["international"] = key_international[1]
            final["international_len"] = len(sorted([(k, v) for k, v in international.items() if v > 0], key=lambda x: x[1]))
        except:
            final["international"] = 999999
            final["international_len"] = 0

        # 国内の分析
        japan = {}
        for sh in sh_dic_japan:
            sh_num = Philosophy.find(sh) + 1
            japan[sh] = Philosophy.find(sh) + 1
        try:
            key_japan = sorted([(k, v) for k, v in japan.items() if v > 0], key=lambda x: x[1])[0]
            final["japan"] = key_japan[1]
            final["japan_len"] = len(sorted([(k, v) for k, v in japan.items() if v > 0], key=lambda x: x[1]))
        except:
            final["japan"] = 999999
            final["japan_len"] = 0

        # それぞれのグループの結果を言及した順にソートする、DataFrameを加工して、出力すること
        # final = sorted(final.items(), key=lambda x: x[1])
        final = [final]
        print(final)
        Item = pd.DataFrame(final)
        Item.to_csv(r"C:\Users\Ray94\Downloads\111.csv", mode='a', index=False, header=None)
