Philosophy = "当社は昭和30年の企業様創業以来、ステンレス鋼の流通を通じてわが国の産業の発展に寄与することを目的とし、販売先と仕入先双方のニーズを調整すると共に、お取引先にソリューションを提供することにより発展してきました。当社の企業理念である「日本一のステンレス・チタン商社として、世のため人のために役立ちたい。」は「ＵＥＸの志」という形にまとめられております。また、この企業理念を具現化すべく経営方針として『ステンレス・チタン商社として価値ある流通機能を提供することで社会に貢献し、永続的な成長を通じてステークホルダー（取引先・社員・株主）の満足度向上をめざします。』を定め、さらなる事業活動の発展に努めるとともに、法令遵守を徹底し、経営体制の一層の強化を目指してまいります。"
sh_dic_customer = ["取引先", "顧客", "お客", "顧客", "クライアント", "消費者", "カスタマー", "需要家", "得意先", "得意様", "ユーザ", "エンドユーザー",
                   "ファン", "企業様", "利用者", "オーナー", "入居者", "不動産所有者", "患者", "施主", "工事業者", "旅行者"]
sh_dic_share = ["株主", "ストックホルダー", "出資者", "投資家"]
sh_dic_staff = ["従業員", "社員", "メンバー", "スタッフ", "クルー"]
sh_dic_community = ["地域社会", "地域", "共同体", "コミュニティ"]

final = {}

customer = {}
for sh in sh_dic_customer:
    sh_num = Philosophy.find(sh) + 1
    customer[sh] = Philosophy.find(sh) + 1
print(customer)
try:
    key_customer = sorted([(k, v) for k, v in customer.items() if v > 0])[0]
    final["customer"] = key_customer[1]
except:
    final["customer"] = 999999


staff = {}
for sh in sh_dic_staff:
    sh_num = Philosophy.find(sh) + 1
    staff[sh] = Philosophy.find(sh) + 1
print(staff)
try:
    key_staff = sorted([(k, v) for k, v in staff.items() if v > 0])[0]
    final["staff"] = key_staff[1]
except:
    final["staff"] = 999999
    
share = {}
for sh in sh_dic_share:
    sh_num = Philosophy.find(sh) + 1
    share[sh] = Philosophy.find(sh) + 1
print(share)
try:
    key_share = sorted([(k, v) for k, v in share.items() if v > 0])[0]
    final["share"] = key_share[1]
except:
    final["share"] = 999999

community = {}
for sh in sh_dic_community:
    sh_num = Philosophy.find(sh) + 1
    community[sh] = Philosophy.find(sh) + 1
print(community)
try:
    key_community = sorted([(k, v) for k, v in community.items() if v > 0])[0]
    final["community"] = key_community[1]
except:
    final["community"] = 999999

final = sorted(final.items(), key=lambda x: x[1])
print(final)

