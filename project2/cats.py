"""5분 성경읽기 고양이 캐릭터 (함께 읽는 친구). 매번 같은 고양이만 나오지 않도록 여섯 마리를 돌아가며 쓴다.

- 성경 인물로 분장·장면 연기 금지. 화면 한쪽에 작게, 꼬리·숨쉬기·눈 깜빡임 정도만 움직임
- 고양이마다 목도리 색을 고정해 시청자가 구별하게 하고, AI 이미지 생성 때는 '기준 설명'을 그대로 붙여 같은 모습을 유지
- color: 시안 그림용 (몸, 무늬, 배·입가, 목도리, 눈)
"""

CATS = [
    dict(name="보리", look="치즈 태비 (주황 줄무늬)", scarf="하늘색 목도리", mood="호기심 많은 막내, 귀를 쫑긋",
         fits="이야기가 많은 본문(창세기·복음서), 섬네일 앞줄",
         prompt="small orange tabby kitten with soft stripes, round amber eyes, light-blue knitted scarf, gentle 2D storybook style",
         color=("#f0a35b", "#d9803a", "#fbe3c8", "#8ec5e8", "#3a2a1a")),
    dict(name="콩이", look="턱시도 (검정·흰색)", scarf="빨간 목도리", mood="점잖고 의젓한 형, 똑바로 앉음",
         fits="차분히 듣는 본문(율법서·서신서)",
         prompt="black-and-white tuxedo cat sitting upright, white chest and paws, green eyes, red knitted scarf, gentle 2D storybook style",
         color=("#2d2f36", "#2d2f36", "#f4f4f2", "#d9534f", "#9bd17a")),
    dict(name="구름", look="하얀 장모", scarf="연분홍 목도리", mood="느긋하고 잘 조는 아이",
         fits="시편·잠언, 묶음 영상 섬네일(잠자며 듣는)",
         prompt="fluffy long-haired white cat with sleepy half-closed blue eyes, pale pink scarf, soft fur, gentle 2D storybook style",
         color=("#f7f7f4", "#e6e4de", "#ffffff", "#f3b6c6", "#6aa6d8")),
    dict(name="모카", look="밀크초코 브라운", scarf="노란 목도리", mood="다정하고 잘 웃는 아이",
         fits="역사서·사도행전",
         prompt="chocolate-brown short-haired cat with cream muzzle, warm honey eyes, mustard-yellow scarf, gentle 2D storybook style",
         color=("#8a5a3c", "#6e4630", "#e9cfb4", "#f2c84b", "#3a2a1a")),
    dict(name="단비", look="회색 단모 (러시안블루 느낌)", scarf="초록 목도리", mood="조용하고 생각 깊은 아이",
         fits="무게 있는 본문(선지서·욥기)",
         prompt="slender blue-gray short-haired cat, calm emerald eyes, forest-green scarf, gentle 2D storybook style",
         color=("#8a94a6", "#77818f", "#c9cfd8", "#4f9a6a", "#5fbf7a")),
    dict(name="호두", look="삼색이 (흰·주황·검정 얼룩)", scarf="보라 목도리", mood="밝고 씩씩한 아이",
         fits="절기 특집, 섬네일 단체 장면",
         prompt="calico cat with white, orange and black patches, bright amber eyes, lavender scarf, gentle 2D storybook style",
         color=("#fbf6ee", "#e58c3c", "#fbf6ee", "#a68bd6", "#3a2a1a")),
]


def assign(segs):
    """업로드 순서(일차·하루 순서)대로 여섯 마리를 돌림. 바로 앞 편과 같은 고양이는 나오지 않고, 편수도 고르게 나뉨."""
    order = sorted(segs, key=lambda s: (s["day"], s["slot"]))
    for k, s in enumerate(order):
        s["cat"] = CATS[k % len(CATS)]
