"""
全問題データを 英検速読トレーナー_DB_v1.xlsx に書き出すスクリプト。
ヘッダー: id, type, eiken_grade, sub_level, passage, japanese_translation,
           question, choice_a, choice_b, choice_c, choice_d, answer,
           target_time_sec, reading_point, evidence_text, reward_point
"""
import openpyxl
import os

data = [
  # ===== 5級 (sub_level=1) =====
  {
    "id": 1, "type": "reading", "eiken_grade": "5", "sub_level": 1,
    "passage": "I have a dog named Kuro. He is black and white. Kuro loves to play in the park. Every weekend, my family goes to the park with Kuro. We throw a ball and Kuro runs to get it. He always comes back wagging his tail. I love playing with Kuro.",
    "japanese_translation": "私はクロという名前の犬を飼っています。彼は白と黒です。クロは公園で遊ぶのが大好きです。毎週末、私の家族はクロと公園に行きます。私たちはボールを投げて、クロはそれを取りに走ります。彼はいつも尻尾を振りながら戻ってきます。私はクロと遊ぶのが大好きです。",
    "question": "Where does the family go every weekend?",
    "choice_a": "To a dog school",
    "choice_b": "To the park",
    "choice_c": "To the beach",
    "choice_d": "To a pet shop",
    "answer": "b",
    "target_time_sec": 25,
    "reading_point": "「Every weekend（毎週末）」という時間表現に注目！習慣的な行動を表しています。その直後に場所が書いてあります。",
    "evidence_text": "Every weekend, my family goes to the park with Kuro",
    "reward_point": 10
  },
  {
    "id": 2, "type": "reading", "eiken_grade": "5", "sub_level": 1,
    "passage": "At school, I eat lunch with my friends. Today I had rice, soup, and fish. The fish was very delicious. My friend Yuki had a sandwich. After lunch, we talked about our favorite foods. I said I like Japanese food the best. Yuki said she likes pizza.",
    "japanese_translation": "学校で私は友達とお昼ご飯を食べます。今日は白米、スープ、魚を食べました。魚はとても美味しかったです。友達のユキはサンドイッチを食べていました。昼食後、私たちは好きな食べ物について話しました。私は和食が一番好きだと言いました。ユキはピザが好きだと言いました。",
    "question": "What did the speaker have for lunch today?",
    "choice_a": "A sandwich",
    "choice_b": "Pizza",
    "choice_c": "Rice, soup, and fish",
    "choice_d": "Noodles",
    "answer": "c",
    "target_time_sec": 25,
    "reading_point": "「Today（今日）」という時間表現の直後に今日の出来事が書かれています。「I had（食べた）」という過去形に注目！",
    "evidence_text": "Today I had rice, soup, and fish",
    "reward_point": 10
  },
  {
    "id": 3, "type": "reading", "eiken_grade": "5", "sub_level": 1,
    "passage": "On Saturday mornings, I always help my mother in the kitchen. We make breakfast together. I wash the vegetables and she cooks them. My father reads the newspaper at the table. After breakfast, we watch TV together as a family. I really enjoy our Saturday mornings.",
    "japanese_translation": "土曜日の朝、私はいつもキッチンでお母さんを手伝います。私たちは一緒に朝食を作ります。私は野菜を洗い、お母さんがそれを料理します。お父さんはテーブルで新聞を読みます。朝食の後、家族みんなでテレビを見ます。私は土曜日の朝がとても好きです。",
    "question": "What does the speaker do to help in the kitchen?",
    "choice_a": "Cooks the vegetables",
    "choice_b": "Reads the newspaper",
    "choice_c": "Makes coffee",
    "choice_d": "Washes the vegetables",
    "answer": "d",
    "target_time_sec": 28,
    "reading_point": "「I wash ... and she cooks」という役割分担の文に注目。主語が「I」と「she」で違うので、それぞれの役割を混同しないように。",
    "evidence_text": "I wash the vegetables and she cooks them",
    "reward_point": 10
  },
  # ===== 4級 (sub_level=1) =====
  {
    "id": 4, "type": "reading", "eiken_grade": "4", "sub_level": 1,
    "passage": "Smartphones have become very common among young people. Many students use smartphones to talk with friends and search for information online. However, some parents worry that children spend too much time on their phones. Doctors say that looking at screens for a long time can hurt your eyes. It is important to take regular breaks from smartphones and spend time outdoors.",
    "japanese_translation": "スマートフォンは若い人の間でとても一般的になりました。多くの生徒がスマートフォンを使って友達と話したり、ネットで情報を調べたりしています。しかし、子供がスマートフォンに時間を使いすぎることを心配している親もいます。医師は、長時間画面を見ると目を傷める可能性があると言っています。スマートフォンから定期的に離れ、屋外で過ごすことが大切です。",
    "question": "Why do some parents worry about their children's smartphone use?",
    "choice_a": "Smartphones are too expensive",
    "choice_b": "Children spend too much time on their phones",
    "choice_c": "Smartphones break easily",
    "choice_d": "Children use phones at school",
    "answer": "b",
    "target_time_sec": 40,
    "reading_point": "「However（しかし）」の後に問題点が書かれることが多い。「worry that（～を心配する）」の後に心配の内容が来ます。",
    "evidence_text": "some parents worry that children spend too much time on their phones",
    "reward_point": 15
  },
  {
    "id": 5, "type": "reading", "eiken_grade": "4", "sub_level": 1,
    "passage": "Last spring, my school class went on a trip to Kyoto. We visited many famous temples and shrines. The most beautiful place was Kinkaku-ji, the Golden Pavilion. We also experienced a traditional Japanese tea ceremony. In the evening, we ate dinner together and shared stories about what we saw. Everyone agreed it was a wonderful trip.",
    "japanese_translation": "先週の春、私たちのクラスは京都に修学旅行に行きました。たくさんの有名な寺社仏閣を訪れました。一番美しかった場所は金閣寺でした。伝統的な日本の茶道も体験しました。夕方には一緒に夕食を食べ、見たものについて話し合いました。みんな素晴らしい旅だったと口をそろえました。",
    "question": "What was the most beautiful place the class visited?",
    "choice_a": "A traditional market",
    "choice_b": "A famous garden",
    "choice_c": "Kinkaku-ji",
    "choice_d": "A mountain shrine",
    "answer": "c",
    "target_time_sec": 40,
    "reading_point": "「The most ～（最も～）」という最上級の表現に注目。これが「一番」を意味します。直後に答えがあります。",
    "evidence_text": "The most beautiful place was Kinkaku-ji, the Golden Pavilion",
    "reward_point": 15
  },
  {
    "id": 6, "type": "reading", "eiken_grade": "4", "sub_level": 1,
    "passage": "Playing sports is very good for your health. When you exercise, your heart becomes stronger and your body produces chemicals that make you feel happy. Many young people enjoy soccer, basketball, and tennis. Even walking for thirty minutes a day can help you stay healthy. The important thing is to keep moving regularly and enjoy the activity you choose.",
    "japanese_translation": "スポーツをすることは健康にとても良いです。運動すると、心臓が強くなり、体内で幸福感をもたらす化学物質が生産されます。多くの若者がサッカー、バスケットボール、テニスを楽しんでいます。1日30分歩くだけでも健康維持に役立ちます。大切なのは、定期的に体を動かし、選んだ活動を楽しむことです。",
    "question": "What happens to your heart when you exercise?",
    "choice_a": "It becomes weaker",
    "choice_b": "It becomes stronger",
    "choice_c": "It beats more slowly",
    "choice_d": "It produces chemicals",
    "answer": "b",
    "target_time_sec": 38,
    "reading_point": "「When you exercise（運動すると）」という条件文の後に、その効果が列挙されています。最初に書かれている効果が答えです。",
    "evidence_text": "When you exercise, your heart becomes stronger",
    "reward_point": 15
  },
  # ===== 3級 (sub_level=1) =====
  {
    "id": 7, "type": "reading", "eiken_grade": "3", "sub_level": 1,
    "passage": "Social media platforms have changed the way young people communicate. Many teenagers use apps to share photos and messages with friends around the world. While social media helps people stay connected, it also has some drawbacks. Some users spend too much time scrolling through posts instead of doing homework or sleeping. Experts recommend limiting social media use to less than two hours a day in order to maintain good mental health.",
    "japanese_translation": "ソーシャルメディアプラットフォームは若者のコミュニケーションの仕方を変えました。多くの10代の若者が、世界中の友達と写真やメッセージを共有するためにアプリを使っています。ソーシャルメディアは人々のつながりを保つ助けになる一方で、欠点もあります。宿題や睡眠の代わりに投稿をスクロールすることに時間を費やしすぎるユーザーもいます。専門家は、良好なメンタルヘルスを維持するために、ソーシャルメディアの使用を1日2時間以内に制限することを勧めています。",
    "question": "What do experts recommend about social media use?",
    "choice_a": "Stopping all social media use",
    "choice_b": "Using only educational apps",
    "choice_c": "Limiting use to less than two hours a day",
    "choice_d": "Sharing only photos, not messages",
    "answer": "c",
    "target_time_sec": 55,
    "reading_point": "「Experts recommend（専門家は勧めている）」という意見表明の表現に注目。その直後に推奨内容が書かれています。",
    "evidence_text": "Experts recommend limiting social media use to less than two hours a day",
    "reward_point": 20
  },
  {
    "id": 8, "type": "reading", "eiken_grade": "3", "sub_level": 1,
    "passage": "Japan has many traditional crafts that have been passed down for hundreds of years. Pottery, weaving, and lacquerware are examples that require years of training to master. Unfortunately, the number of craftspeople is decreasing because young people tend to choose modern careers instead. Some local governments have started programs to teach traditional skills to younger generations, hoping to keep these crafts alive. However, it remains difficult to earn a living solely from traditional crafts in today's competitive economy.",
    "japanese_translation": "日本には数百年にわたって受け継がれてきた伝統工芸品がたくさんあります。陶芸、機織り、漆器などは、習得するのに何年もの訓練が必要な例です。残念ながら、若者が現代的な職業を選ぶ傾向にあるため、職人の数は減っています。伝統的な技術を若い世代に教えるプログラムを始めた地方自治体もあります。しかし、今日の競争的な経済の中で伝統工芸だけで生計を立てることは依然として難しいままです。",
    "question": "Why is the number of craftspeople decreasing?",
    "choice_a": "Traditional crafts have become unpopular globally",
    "choice_b": "The government stopped supporting craftspeople",
    "choice_c": "Young people tend to choose modern careers",
    "choice_d": "Craft materials are too expensive",
    "answer": "c",
    "target_time_sec": 55,
    "reading_point": "「because（なぜなら）」という接続詞の後に理由が書かれています。問題文の「Why（なぜ）」に対応する原因を素早く探しましょう。",
    "evidence_text": "young people tend to choose modern careers instead",
    "reward_point": 20
  },
  {
    "id": 9, "type": "reading", "eiken_grade": "3", "sub_level": 1,
    "passage": "Food waste is a serious environmental problem in many countries. In Japan, millions of tons of food are thrown away every year, even though much of it is still safe to eat. Restaurants and supermarkets sometimes discard food approaching its expiration date. Overpurchasing and strict standards for food appearance contribute to the problem. Individuals can help by planning meals carefully, buying only what is needed, and using leftovers creatively before they spoil.",
    "japanese_translation": "食品廃棄物は多くの国で深刻な環境問題です。日本では、まだ食べられるにもかかわらず、毎年何百万トンもの食品が廃棄されています。レストランやスーパーマーケットは、賞味期限が近い食品を廃棄することがあります。買いすぎや食品の見た目に関する厳しい基準が問題を深刻にしています。食事を丁寧に計画し、必要なものだけを買い、傷む前に残り物を工夫して使うことで、個人も貢献できます。",
    "question": "What is one way individuals can reduce food waste?",
    "choice_a": "Buying food in large quantities to save money",
    "choice_b": "Eating at restaurants more often",
    "choice_c": "Planning meals carefully",
    "choice_d": "Avoiding all processed food",
    "answer": "c",
    "target_time_sec": 55,
    "reading_point": "最後の文に「Individuals can help by（個人は～することで貢献できる）」という形で解決策が3つ列挙されています。最初に挙げられているものが答えです。",
    "evidence_text": "Individuals can help by planning meals carefully",
    "reward_point": 20
  },
  # ===== 準2 (sub_level=1) =====
  {
    "id": 10, "type": "reading", "eiken_grade": "pre2", "sub_level": 1,
    "passage": "E-sports, or competitive video gaming, has grown into a billion-dollar industry in recent years. Professional players can earn large salaries and compete in tournaments held in major sports arenas worldwide. Unlike traditional sports, e-sports does not require physical strength, making it accessible to a wider range of participants. However, critics argue that playing video games for many hours daily can lead to health problems such as poor posture and insufficient physical exercise. Despite such concerns, many universities now offer e-sports scholarships, treating competitive gaming as a legitimate athletic pursuit.",
    "japanese_translation": "eスポーツ、つまり競技ビデオゲームは、近年十億ドル規模の産業に成長しました。プロ選手は高い給料を得て、世界中の大型スポーツアリーナで開催されるトーナメントに出場することができます。従来のスポーツとは異なり、eスポーツは身体的な強さを必要としないため、より幅広い参加者が参加できます。しかし批評家たちは、毎日何時間もビデオゲームをすることで、姿勢の悪化や運動不足などの健康問題を引き起こす可能性があると主張しています。そのような懸念があるにもかかわらず、多くの大学が今やeスポーツの奨学金を提供し、競技ゲームを正当なスポーツ活動として扱っています。",
    "question": "What advantage does e-sports have over traditional sports?",
    "choice_a": "It pays higher salaries than other sports",
    "choice_b": "It does not require physical strength",
    "choice_c": "It is played outdoors in all weather",
    "choice_d": "It is cheaper to participate in",
    "answer": "b",
    "target_time_sec": 70,
    "reading_point": "「Unlike（～とは異なり）」という比較の表現に注目。この後にeスポーツの特徴（＝従来のスポーツとの違い）が書かれています。",
    "evidence_text": "Unlike traditional sports, e-sports does not require physical strength, making it accessible to a wider range of participants",
    "reward_point": 25
  },
  {
    "id": 11, "type": "reading", "eiken_grade": "pre2", "sub_level": 1,
    "passage": "As concerns about climate change grow, many countries are investing heavily in renewable energy sources such as solar and wind power. These technologies generate electricity without releasing greenhouse gases that contribute to global warming. One significant challenge is reliability — solar panels produce less power on cloudy days, and wind turbines require sufficient wind. Scientists are actively developing improved battery systems to store energy for periods when generation is low. If these storage challenges can be overcome, renewable energy could eventually replace fossil fuels as the world's primary energy source.",
    "japanese_translation": "気候変動への懸念が高まる中、多くの国が太陽光や風力などの再生可能エネルギー源に多額の投資をしています。これらの技術は、地球温暖化の原因となる温室効果ガスを排出せずに電力を生産します。大きな課題の一つは信頼性です。太陽光パネルは曇りの日には発電量が減り、風力タービンは十分な風が必要です。科学者たちは、発電量が少ない時期のためにエネルギーを蓄える改良された蓄電システムの開発に取り組んでいます。これらの蓄電の課題が克服されれば、再生可能エネルギーは最終的に世界の主要エネルギー源として化石燃料に取って代わることができるかもしれません。",
    "question": "What is described as a significant challenge for renewable energy?",
    "choice_a": "The high cost of manufacturing panels",
    "choice_b": "The lack of government support",
    "choice_c": "Its unreliable output depending on weather",
    "choice_d": "The shortage of trained engineers",
    "answer": "c",
    "target_time_sec": 72,
    "reading_point": "「One significant challenge is（大きな課題の一つは）」という導入フレーズの後に、その内容が具体的に説明されています。ダッシュ（—）の後に解説が続きます。",
    "evidence_text": "solar panels produce less power on cloudy days, and wind turbines require sufficient wind",
    "reward_point": 25
  },
  {
    "id": 12, "type": "reading", "eiken_grade": "pre2", "sub_level": 1,
    "passage": "The COVID-19 pandemic caused millions of workers worldwide to begin working from home. Many companies discovered that remote work can be equally as productive as office work. Employees benefit from eliminating commute time and gaining greater flexibility in managing their schedules. However, working from home presents its own challenges. Some employees feel isolated from colleagues and struggle to maintain the boundary between professional and personal time. As a result, many organizations are now experimenting with hybrid models, allowing staff to split their time between home and office.",
    "japanese_translation": "COVID-19パンデミックにより、世界中で何百万人もの労働者が在宅勤務を始めました。多くの企業は、リモートワークがオフィスでの仕事と同じくらい生産的であることを発見しました。従業員は通勤時間をなくし、スケジュール管理においてより大きな柔軟性を得るというメリットがあります。しかし在宅勤務には独自の課題もあります。同僚から孤立感を感じたり、仕事と個人の時間の境界を維持するのが難しいと感じる従業員もいます。その結果、多くの組織が今、スタッフが自宅とオフィスで時間を分けられるハイブリッドモデルを試験的に導入しています。",
    "question": "What is one challenge of working from home mentioned in the passage?",
    "choice_a": "Employees work longer hours",
    "choice_b": "Technology costs increase",
    "choice_c": "Some employees feel isolated from colleagues",
    "choice_d": "Managers cannot supervise workers",
    "answer": "c",
    "target_time_sec": 68,
    "reading_point": "「However（しかし）」の後にデメリットが来るパターン。「presents its own challenges（独自の課題がある）」の後に具体的な課題が続きます。",
    "evidence_text": "Some employees feel isolated from colleagues",
    "reward_point": 25
  },
  # ===== 準2+ (pre2 sub_level=2) =====
  {
    "id": 13, "type": "reading", "eiken_grade": "pre2", "sub_level": 2,
    "passage": "Artificial intelligence is increasingly being integrated into educational settings to personalize learning for individual students. AI-powered platforms analyze performance data to identify each student's weaknesses and automatically generate customized practice materials. Supporters argue this approach is more efficient than traditional classroom instruction, where teachers must simultaneously address the diverse needs of an entire class. However, critics raise significant concerns regarding student privacy, as these systems continuously collect and process large volumes of personal learning data. Furthermore, some educators worry that excessive dependence on AI tools may limit opportunities for students to develop independent critical thinking through human interaction.",
    "japanese_translation": "人工知能は、個々の生徒に合わせた学習をパーソナライズするために、教育現場でますます活用されています。AIを搭載したプラットフォームは、学習データを分析して各生徒の弱点を特定し、自動的にカスタマイズされた練習教材を生成します。支持者たちは、このアプローチは、教師がクラス全体の多様なニーズに同時に対応しなければならない従来の授業よりも効率的だと主張しています。しかし批評家たちは、これらのシステムが大量の個人学習データを継続的に収集・処理するため、生徒のプライバシーに関する重大な懸念を提起しています。さらに、一部の教育者は、AIツールへの過度な依存が、人間との交流を通じた自立した批判的思考を発達させる機会を制限する可能性があると心配しています。",
    "question": "What concern do critics raise about AI in education?",
    "choice_a": "AI tools are too expensive for schools",
    "choice_b": "AI cannot replace human teachers",
    "choice_c": "Student privacy may be compromised",
    "choice_d": "AI makes learning less enjoyable",
    "answer": "c",
    "target_time_sec": 82,
    "reading_point": "「However, critics raise ... concerns regarding（しかし批評家は～への懸念を）」という譲歩→批判の流れに注目。「regarding（～に関して）」の後が懸念の内容です。",
    "evidence_text": "critics raise significant concerns regarding student privacy, as these systems continuously collect and process large volumes of personal learning data",
    "reward_point": 28
  },
  {
    "id": 14, "type": "reading", "eiken_grade": "pre2", "sub_level": 2,
    "passage": "As cities worldwide expand at unprecedented rates, urban planners face the complex challenge of balancing development with the preservation of green spaces. Parks and urban forests provide critical environmental services, including filtering air pollutants and mitigating the urban heat island effect, whereby cities become significantly warmer than surrounding rural areas. Research consistently demonstrates that access to green spaces positively influences the mental health and overall well-being of urban residents. Nevertheless, rapidly rising land values in growing cities make it financially challenging to reserve open land for parks. Innovative urban designers are responding by incorporating greenery directly into architecture through rooftop gardens, green walls, and elevated park structures.",
    "japanese_translation": "世界中の都市が前例のない速さで拡大するにつれて、都市計画者は開発と緑地の保全を両立させるという複雑な課題に直面しています。公園や都市の森林は、大気汚染物質を濾過したり、都市が周辺の農村地域よりも著しく暑くなる「ヒートアイランド現象」を緩和したりするなど、重要な環境サービスを提供しています。研究は一貫して、緑地へのアクセスが都市住民の精神的健康と全体的な幸福度にプラスの影響を与えることを示しています。それにもかかわらず、成長する都市では急速に上昇する地価が、公園として土地を確保することを財政的に難しくしています。革新的な都市デザイナーたちは、屋上庭園、グリーンウォール、高架公園構造物を通じて、建築物に直接緑を組み込むことで対応しています。",
    "question": "What is the 'urban heat island effect'?",
    "choice_a": "A type of air pollution found only in cities",
    "choice_b": "The phenomenon where cities become warmer than rural areas",
    "choice_c": "The loss of green spaces due to construction",
    "choice_d": "A design method for cooling urban buildings",
    "answer": "b",
    "target_time_sec": 85,
    "reading_point": "「whereby（それによって）」という関係副詞は直前の名詞の説明を続けます。「urban heat island effect」の直後にカンマ＋wherebyで定義が書かれています。",
    "evidence_text": "cities become significantly warmer than surrounding rural areas",
    "reward_point": 28
  },
  {
    "id": 15, "type": "reading", "eiken_grade": "pre2", "sub_level": 2,
    "passage": "The fast fashion industry, characterized by rapid production of large volumes of cheaply made clothing, has attracted mounting criticism for its severe environmental impact. The manufacture of synthetic fabrics consumes vast quantities of water and generates substantial carbon emissions throughout the production process. Because fast fashion garments are designed for disposability rather than durability, they are discarded after minimal use, contributing significantly to landfill waste. In response to growing consumer pressure, some fashion brands have launched sustainable clothing lines using recycled or organic materials. Critics, however, note that these initiatives represent only a fraction of overall production and that fundamental changes to the industry's business model are required to achieve meaningful environmental progress.",
    "japanese_translation": "安価な衣服を大量かつ迅速に生産することを特徴とするファストファッション産業は、深刻な環境への影響により批判が高まっています。合成繊維の製造は膨大な量の水を消費し、製造プロセス全体で大量の炭素排出を生み出します。ファストファッションの衣服は耐久性よりも使い捨てを前提に設計されているため、最小限の使用後に廃棄され、埋め立てごみに大きく貢献しています。消費者からの圧力の高まりに応じて、一部のファッションブランドがリサイクルまたはオーガニック素材を使用したサステナブルな衣類ラインを発売しています。しかし批評家たちは、これらの取り組みは全体の生産のほんの一部に過ぎず、意味のある環境改善を達成するためには産業のビジネスモデルの根本的な変革が必要だと指摘しています。",
    "question": "Why do fast fashion items contribute to landfill waste?",
    "choice_a": "They are made from materials that cannot be recycled",
    "choice_b": "They are designed for disposability rather than durability",
    "choice_c": "Consumers choose to throw them away for fashion reasons",
    "choice_d": "They deteriorate quickly due to poor storage",
    "answer": "b",
    "target_time_sec": 88,
    "reading_point": "「Because（なぜなら）」の後に原因が書かれています。「designed for X rather than Y（YよりもXのために設計されている）」という比較表現がポイントです。",
    "evidence_text": "fast fashion garments are designed for disposability rather than durability",
    "reward_point": 28
  },
  # ===== 2級 (sub_level=1) =====
  {
    "id": 16, "type": "reading", "eiken_grade": "2", "sub_level": 1,
    "passage": "The placebo effect refers to the measurable improvement in a patient's condition following a treatment that contains no active medical ingredient. Research has demonstrated that when patients believe they are receiving effective treatment, the brain releases chemicals such as endorphins that genuinely alleviate symptoms. This phenomenon has significant implications for medicine, suggesting that a patient's psychological state plays a crucial role in the recovery process. However, administering placebos in clinical settings raises ethical concerns, as it involves deliberately deceiving patients about the nature of their treatment. Some researchers have therefore advocated for 'open-label placebos,' in which patients are explicitly informed that they are receiving an inert treatment, yet still experience measurable improvements.",
    "japanese_translation": "プラセボ効果とは、有効な医薬成分を含まない治療を受けた後に患者の状態が測定可能なほど改善することを指します。研究によると、患者が効果的な治療を受けていると信じると、脳がエンドルフィンなどの化学物質を放出し、実際に症状を緩和することが示されています。この現象は医療に重要な意味を持ち、患者の心理状態が回復過程で重要な役割を果たすことを示唆しています。しかし、臨床環境でプラセボを投与することは、患者に治療の性質について故意に欺くことを伴うため、倫理的な懸念を生じさせます。そこで一部の研究者は、「オープンラベルプラセボ」を提唱しています。これは、患者が不活性な治療を受けていることを明示的に知らされながらも、測定可能な改善を経験するというものです。",
    "question": "What ethical concern does the use of placebos in clinical settings raise?",
    "choice_a": "Placebos are more expensive than real medicine",
    "choice_b": "It involves deliberately deceiving patients",
    "choice_c": "Placebos can cause harmful side effects",
    "choice_d": "It undermines the authority of doctors",
    "answer": "b",
    "target_time_sec": 90,
    "reading_point": "「raises ethical concerns（倫理的懸念を生じさせる）」という表現の後に「as it involves（それが～を伴うため）」という理由節が続いています。この接続が答えへの鍵です。",
    "evidence_text": "it involves deliberately deceiving patients about the nature of their treatment",
    "reward_point": 30
  },
  {
    "id": 17, "type": "reading", "eiken_grade": "2", "sub_level": 1,
    "passage": "Traditional economics operates on the assumption that individuals make rational decisions based on complete information and consistent self-interest. However, the emerging field of behavioral economics, pioneered by researchers including Daniel Kahneman and Richard Thaler, challenges this assumption by documenting how cognitive biases and emotional influences systematically distort human judgment. A particularly well-documented phenomenon is loss aversion — the tendency to experience the pain of a loss as approximately twice as powerful as the pleasure derived from an equivalent gain. Drawing on these insights, policymakers and businesses have increasingly employed 'nudge' strategies, designing environments that steer people toward beneficial choices without restricting freedom, such as automatically enrolling employees in retirement savings plans.",
    "japanese_translation": "伝統的な経済学は、個人が完全な情報と一貫した自己利益に基づいて合理的な決定を下すという前提のもとに機能しています。しかし、ダニエル・カーネマンやリチャード・セイラーなどの研究者が先駆者となった行動経済学という新興分野は、認知バイアスと感情的影響が人間の判断を体系的に歪めることを実証することで、この前提に疑問を呈しています。特によく記録されている現象は損失回避性です。これは、損失の痛みを同等の利益から得られる喜びの約2倍強く感じる傾向のことです。これらの知見を活かして、政策立案者や企業はますます「ナッジ」戦略を採用しています。これは、従業員を退職貯蓄プランに自動的に加入させるなど、自由を制限せずに人々を有益な選択に誘導するよう環境を設計することです。",
    "question": "What is 'loss aversion' as described in the passage?",
    "choice_a": "The preference for smaller but certain rewards",
    "choice_b": "The tendency to feel losses more strongly than equivalent gains",
    "choice_c": "The avoidance of all financial risk",
    "choice_d": "The bias toward familiar options over new ones",
    "answer": "b",
    "target_time_sec": 95,
    "reading_point": "「A particularly well-documented phenomenon is（特によく記録されている現象は）」という強調表現の後にダッシュで定義が続きます。定義文の構造を素早く捉えましょう。",
    "evidence_text": "the tendency to experience the pain of a loss as approximately twice as powerful as the pleasure derived from an equivalent gain",
    "reward_point": 30
  },
  {
    "id": 18, "type": "reading", "eiken_grade": "2", "sub_level": 1,
    "passage": "Despite covering over seventy percent of Earth's surface, the world's oceans remain substantially unexplored. The extreme conditions characteristic of the deep sea — including crushing pressure, near-freezing temperatures, and total absence of sunlight — make exploration technically demanding and financially prohibitive. Recent advances in remotely operated vehicles and autonomous underwater vehicles have enabled scientists to survey previously inaccessible regions of the seafloor and document extraordinary ecosystems thriving around hydrothermal vents, where organisms survive without photosynthesis. These discoveries have fundamentally revised scientific understanding of the conditions under which life can exist and carry implications for the search for extraterrestrial life. Nevertheless, researchers estimate that less than twenty percent of the ocean floor has been systematically mapped to date.",
    "japanese_translation": "地球の表面の70パーセント以上を覆っているにもかかわらず、世界の海洋は依然として大部分が探索されていません。深海に特徴的な極端な条件、つまり押しつぶされるような圧力、氷点近くの温度、そして太陽光の完全な不在が、探索を技術的に困難で財政的に費用のかかるものにしています。遠隔操作車両や自律型水中ビークルの最近の進歩により、科学者たちはこれまで到達不可能だった海底の地域を調査し、光合成なしに生存する生物が生息する熱水噴出孔周辺の驚くべき生態系を記録することが可能になりました。これらの発見は、生命が存在できる条件に関する科学的理解を根本的に改訂し、地球外生命体の探索にも意味を持ちます。それにもかかわらず、研究者たちは海底の20パーセント未満が今日まで体系的にマッピングされていると推定しています。",
    "question": "What have recent advances in underwater vehicles enabled scientists to discover?",
    "choice_a": "New mineral deposits beneath the ocean floor",
    "choice_b": "Ecosystems thriving around hydrothermal vents without sunlight",
    "choice_c": "Evidence of ancient civilizations on the seafloor",
    "choice_d": "The precise depth of the deepest ocean trenches",
    "answer": "b",
    "target_time_sec": 95,
    "reading_point": "「have enabled scientists to（科学者たちが～することを可能にした）」という結果の表現の後に2つの成果が「and」で結ばれています。問題の「discover」に対応する内容を探しましょう。",
    "evidence_text": "extraordinary ecosystems thriving around hydrothermal vents, where organisms survive without photosynthesis",
    "reward_point": 30
  },
  # ===== 準1 (sub_level=1) =====
  {
    "id": 19, "type": "reading", "eiken_grade": "pre1", "sub_level": 1,
    "passage": "Cognitive load theory, formulated by educational psychologist John Sweller, posits that the human working memory — the cognitive system responsible for temporarily holding and manipulating information — operates under strict capacity constraints. When learners are exposed to instructional materials that impose excessive cognitive demands, the available mental resources become overwhelmed, and meaningful learning is impeded. Sweller distinguished three constituent types of cognitive load: intrinsic load, which reflects the inherent complexity of the subject matter; extraneous load, arising from unnecessary or poorly structured instructional design; and germane load, representing the cognitive effort invested in constructing and consolidating durable knowledge schemas. Effective pedagogical design seeks to minimize extraneous load — thereby freeing cognitive resources — while simultaneously optimizing germane load to promote deep, transferable understanding.",
    "japanese_translation": "教育心理学者ジョン・スウェラーによって定式化された認知負荷理論は、情報を一時的に保持・操作する認知システムである人間のワーキングメモリが、厳しい容量制限のもとで機能することを主張しています。学習者が過度な認知的要求を課す教材にさらされると、利用可能な精神的資源が圧倒され、意味のある学習が妨げられます。スウェラーは認知負荷を3つの構成要素に区別しました。内在的負荷（学習内容そのものの複雑さを反映するもの）、外在的負荷（不必要または構造が不十分な教授法のデザインから生じるもの）、そして学習関連負荷（耐久性のある知識スキーマを構築・強化するために投じられる認知的努力を表すもの）です。効果的な教授法のデザインは、認知リソースを解放するために外在的負荷を最小化しながら、同時に深く応用可能な理解を促進するために学習関連負荷を最適化することを目指します。",
    "question": "According to Sweller, what is 'extraneous load'?",
    "choice_a": "The complexity inherent in the subject being taught",
    "choice_b": "The mental effort used to build long-term knowledge",
    "choice_c": "Cognitive load caused by poor instructional design",
    "choice_d": "The total capacity of working memory",
    "answer": "c",
    "target_time_sec": 110,
    "reading_point": "3種類の認知負荷が列挙されている段落の構造を把握しましょう。コロン（:）の後に3つが順に説明されます。各タイプの定義を説明する関係詞節（which/arising from/representing）に注目。",
    "evidence_text": "extraneous load, arising from unnecessary or poorly structured instructional design",
    "reward_point": 35
  },
  {
    "id": 20, "type": "reading", "eiken_grade": "pre1", "sub_level": 1,
    "passage": "The tragedy of the commons, a concept introduced by ecologist Garrett Hardin in a seminal 1968 essay, describes the dynamic by which rational individuals, each pursuing personal benefit, collectively degrade a finite shared resource to the detriment of the entire community. Hardin illustrated this with the analogy of herders sharing open pastureland: each individual has a rational incentive to graze additional livestock to maximize personal gain, but if all act similarly, the commons becomes irreversibly degraded. Hardin's framework has been extensively applied to contemporary challenges including overfishing, atmospheric carbon accumulation, and groundwater depletion. However, economist Elinor Ostrom, awarded the Nobel Prize in Economics in 2009, empirically challenged Hardin's pessimistic determinism by documenting numerous cases in which communities have successfully developed and enforced self-governing institutions for the sustainable management of shared resources.",
    "japanese_translation": "「コモンズの悲劇」は、生態学者ギャレット・ハーディンが1968年の画期的な論文で提唱した概念で、各自が個人的な利益を追求する合理的な個人たちが、コミュニティ全体の不利益となるように有限の共有資源を集合的に劣化させるダイナミクスを説明しています。ハーディンは、開放された牧草地を共有する牧夫たちの比喩でこれを説明しました。各個人には個人的な利益を最大化するために追加の家畜を放牧する合理的な動機がありますが、全員が同様に行動すれば、コモンズは取り返しのつかないほど劣化してしまいます。ハーディンの枠組みは、乱獲、大気中の炭素蓄積、地下水の枯渇などの現代的な課題に広く応用されています。しかし、2009年にノーベル経済学賞を受賞した経済学者エリノア・オストロムは、コミュニティが共有資源の持続可能な管理のための自治機構を成功裏に開発・施行した数多くの事例を記録することで、ハーディンの悲観的決定論に実証的に挑戦しました。",
    "question": "How did Elinor Ostrom challenge Hardin's theory?",
    "choice_a": "By proving that shared resources cannot be sustainably managed",
    "choice_b": "By showing communities can self-govern shared resources sustainably",
    "choice_c": "By demonstrating that private ownership prevents resource depletion",
    "choice_d": "By arguing that government regulation is always necessary",
    "answer": "b",
    "target_time_sec": 115,
    "reading_point": "「However（しかし）」の後にオストロムの反論が続きます。「by documenting（記録することによって）」という手段の表現の後に、何を示したかが書かれています。長い文を主節と従属節に分けて読みましょう。",
    "evidence_text": "communities have successfully developed and enforced self-governing institutions for the sustainable management of shared resources",
    "reward_point": 35
  },
  {
    "id": 21, "type": "reading", "eiken_grade": "pre1", "sub_level": 1,
    "passage": "Neuroplasticity — the capacity of the brain to reorganize its structural and functional architecture through the formation of new synaptic connections and the pruning of unused pathways — has profoundly unsettled the previously dominant view that the adult brain is essentially static. Longitudinal neuroimaging studies have demonstrated that sustained engagement in cognitively demanding activities, acquisition of novel skills, and aerobic exercise can all stimulate neurogenesis and reinforce synaptic efficiency. These findings carry far-reaching implications for education, neurorehabilitation, and our conceptual understanding of how behavioral patterns are acquired and extinguished. Crucially, however, neuroplasticity is not inherently adaptive: the identical mechanisms that facilitate skill acquisition and recovery from injury also underlie the entrenchment of compulsive behaviors and the consolidation of trauma-associated neural circuits. The central challenge for contemporary neuroscience therefore lies in delineating the specific conditions and interventions that reliably promote adaptive over maladaptive neuroplastic change.",
    "japanese_translation": "神経可塑性とは、新しいシナプス結合の形成と使われていない経路の刈り込みを通じて脳が構造的・機能的アーキテクチャを再編成する能力のことで、成人の脳は本質的に静的であるという従来の支配的な見方を根本的に覆しました。縦断的な神経画像研究により、認知的に要求の高い活動への持続的な取り組み、新しいスキルの習得、有酸素運動がすべて神経新生を刺激しシナプス効率を強化できることが示されています。これらの発見は、教育、神経リハビリテーション、そして行動パターンがどのように習得・消去されるかについての概念的理解に広範な意味を持ちます。しかし重要なことに、神経可塑性は本質的に適応的なわけではありません。スキル習得や損傷からの回復を促進するのと全く同じメカニズムが、強迫行動の定着やトラウマ関連神経回路の固化の根底にもあります。したがって、現代神経科学の核心的な課題は、適応的な神経可塑的変化を不適応的なものよりも確実に促進する特定の条件と介入を明確にすることにあります。",
    "question": "What crucial limitation of neuroplasticity does the passage highlight?",
    "choice_a": "It declines sharply after the age of thirty",
    "choice_b": "It cannot occur without professional medical intervention",
    "choice_c": "The same mechanisms can also reinforce harmful patterns",
    "choice_d": "It only operates during periods of intense physical exercise",
    "answer": "c",
    "target_time_sec": 120,
    "reading_point": "「Crucially, however（しかし重要なことに）」という強い逆接の副詞に注目。この後に神経可塑性の「落とし穴」が書かれています。「the identical mechanisms that（同じメカニズムが）」という同格表現で両面が説明されています。",
    "evidence_text": "the identical mechanisms that facilitate skill acquisition and recovery from injury also underlie the entrenchment of compulsive behaviors",
    "reward_point": 35
  }
]

HEADERS = [
    "id", "type", "eiken_grade", "sub_level", "passage", "japanese_translation",
    "question", "choice_a", "choice_b", "choice_c", "choice_d", "answer",
    "target_time_sec", "reading_point", "evidence_text", "reward_point"
]

OUTPUT = "英検速読トレーナー_DB_v1.xlsx"

script_dir = os.path.dirname(os.path.abspath(__file__))
output_path = os.path.join(script_dir, OUTPUT)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "passages"

# ヘッダー行
ws.append(HEADERS)

# データ行
for d in data:
    row = [d.get(h, "") for h in HEADERS]
    ws.append(row)

wb.save(output_path)
print(f"完了！ {OUTPUT} に {len(data)} 問を書き出しました。")
print(f"保存先: {output_path}")
