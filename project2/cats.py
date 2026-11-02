"""5분 성경읽기 고양이 캐릭터 (함께 읽는 친구). 매번 같은 고양이만 나오지 않도록 열두 마리를 돌아가며 쓴다.

- 성경 인물로 분장·장면 연기 금지. 화면 한쪽에 작게, 꼬리·숨쉬기·눈 깜빡임 정도만 움직임
- 고양이마다 목도리 색을 고정해 시청자가 구별하게 하고, AI 이미지 생성 때는 '기준 설명'을 그대로 붙여 같은 모습을 유지
- color: 시안 그림용 (몸, 무늬, 배·입가, 목도리, 눈), pattern: 시안 그림 무늬 종류
"""

CATS = [
    dict(name="보리", look="치즈 태비 (주황 줄무늬)", scarf="하늘색 목도리", mood="호기심 많은 막내, 귀를 쫑긋",
         fits="이야기가 많은 본문(창세기·복음서), 섬네일 앞줄",
         prompt="small orange tabby kitten with soft stripes, round amber eyes, light-blue knitted scarf, gentle 2D storybook style",
         color=("#f0a35b", "#d9803a", "#fbe3c8", "#8ec5e8", "#3a2a1a"), pattern="tabby"),
    dict(name="콩이", look="턱시도 (검정·흰색)", scarf="빨간 목도리", mood="점잖고 의젓한 형, 똑바로 앉음",
         fits="차분히 듣는 본문(율법서·서신서)",
         prompt="black-and-white tuxedo cat sitting upright, white chest and paws, green eyes, red knitted scarf, gentle 2D storybook style",
         color=("#2d2f36", "#2d2f36", "#f4f4f2", "#d9534f", "#9bd17a"), pattern="tuxedo"),
    dict(name="구름", look="하얀 장모", scarf="연분홍 목도리", mood="느긋하고 잘 조는 아이",
         fits="시편·잠언, 묶음 영상 섬네일(잠자며 듣는)",
         prompt="fluffy long-haired white cat with sleepy half-closed blue eyes, pale pink scarf, soft fur, gentle 2D storybook style",
         color=("#f7f7f4", "#e6e4de", "#ffffff", "#f3b6c6", "#6aa6d8"), pattern="long"),
    dict(name="모카", look="밀크초코 브라운", scarf="노란 목도리", mood="다정하고 잘 웃는 아이",
         fits="역사서·사도행전",
         prompt="chocolate-brown short-haired cat with cream muzzle, warm honey eyes, mustard-yellow scarf, gentle 2D storybook style",
         color=("#8a5a3c", "#6e4630", "#e9cfb4", "#f2c84b", "#3a2a1a"), pattern="solid"),
    dict(name="단비", look="회색 단모 (러시안블루 느낌)", scarf="초록 목도리", mood="조용하고 생각 깊은 아이",
         fits="무게 있는 본문(선지서·욥기)",
         prompt="slender blue-gray short-haired cat, calm emerald eyes, forest-green scarf, gentle 2D storybook style",
         color=("#8a94a6", "#77818f", "#c9cfd8", "#4f9a6a", "#5fbf7a"), pattern="solid"),
    dict(name="호두", look="삼색이 (흰·주황·검정 얼룩)", scarf="보라 목도리", mood="밝고 씩씩한 아이",
         fits="절기 특집, 섬네일 단체 장면",
         prompt="calico cat with white, orange and black patches, bright amber eyes, lavender scarf, gentle 2D storybook style",
         color=("#fbf6ee", "#e58c3c", "#fbf6ee", "#a68bd6", "#3a2a1a"), pattern="calico"),
    dict(name="누리", look="고등어 태비 (회갈색 줄무늬)", scarf="주황 목도리", mood="부지런하고 듬직한 아이",
         fits="역사서·여호수아처럼 길을 가는 본문",
         prompt="gray-brown mackerel tabby cat with bold stripes, olive-green eyes, orange knitted scarf, gentle 2D storybook style",
         color=("#9a8f80", "#5f5649", "#e6ddd0", "#f08a3c", "#8aa04a"), pattern="tabby"),
    dict(name="별이", look="까만 단모", scarf="민트 목도리", mood="수줍지만 반짝이는 눈",
         fits="밤 장면, 잠들기 전 묶음 섬네일",
         prompt="sleek all-black short-haired cat with big golden-yellow eyes, mint-green scarf, gentle 2D storybook style",
         color=("#26262c", "#26262c", "#3a3a42", "#7fd6c0", "#f2c84b"), pattern="solid"),
    dict(name="솜이", look="크림색 접힌 귀 (스코티시폴드 느낌)", scarf="와인색 목도리", mood="동글동글 순한 아이",
         fits="시편·아가처럼 부드러운 본문",
         prompt="round cream-colored cat with small folded ears, round copper eyes, wine-red scarf, gentle 2D storybook style",
         color=("#f1dfbf", "#e2c99c", "#fbf2e2", "#8e2f46", "#c9803a"), pattern="fold"),
    dict(name="율무", look="샴 (크림 몸·갈색 얼굴과 귀)", scarf="남색 목도리", mood="말이 많은 수다쟁이",
         fits="서신서·대화가 많은 본문",
         prompt="siamese cat with cream body and dark brown face mask, ears, paws and tail, bright blue eyes, navy-blue scarf, gentle 2D storybook style",
         color=("#efe3cf", "#5a4033", "#f7efe2", "#2f3f7a", "#5aa6e0"), pattern="point"),
    dict(name="망고", look="치즈 바이컬러 (주황·흰색)", scarf="갈색 목도리", mood="장난기 많은 개구쟁이 (무거운 본문엔 얌전히)",
         fits="복음서 기적 이야기, 쇼츠 섬네일",
         prompt="orange-and-white bicolor cat with white muzzle and chest, amber eyes, brown knitted scarf, gentle 2D storybook style",
         color=("#f2a14d", "#f2a14d", "#fffaf2", "#8a5a3c", "#3a2a1a"), pattern="tuxedo"),
    dict(name="라떼", look="회색·흰색 바이컬러", scarf="검정 목도리", mood="어른스러운 맏언니",
         fits="율법서·선지서처럼 차분한 본문",
         prompt="gray-and-white bicolor cat with white face blaze and chest, gentle hazel eyes, black knitted scarf, gentle 2D storybook style",
         color=("#9aa0aa", "#9aa0aa", "#fbfbfa", "#2a2a30", "#b9a05a"), pattern="tuxedo"),
]


def assign(segs):
    """업로드 순서(일차·하루 순서)대로 열두 마리를 돌림. 바로 앞 편과 같은 고양이는 나오지 않고, 편수도 고르게 나뉨."""
    order = sorted(segs, key=lambda s: (s["day"], s["slot"]))
    for k, s in enumerate(order):
        s["cat"] = CATS[k % len(CATS)]
