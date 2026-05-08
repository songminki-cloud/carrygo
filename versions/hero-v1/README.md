# CarryGo Homepage Hero v1

Saved on 2026-05-08 before starting v2 hero work.

## Purpose
This folder preserves the finalized v1 hero state so it can be restored anytime.

## v1 hero structure
- No separate hero illustration/image asset.
- Hero is typography-led text inside `index.html`.
- Logo asset used in header: `assets/logo/carrygo_logo_final.png`

## v1 hero text

### English typographic headline
```text
Drop
luggage.
Enjoy
the show.
PICK IT UP!
```

### Korean supporting copy
```text
CarryGo는 공연 관객을 위한 짐 보관 서비스입니다.
무거운 짐은 맡기고 가볍게 입장하세요.
```

### CTA buttons
```text
짐 보관 신청하기
이용 방법 보기
```

## Restore guide
To restore the full v1 homepage file:

```bash
cp versions/hero-v1/index.html index.html
```

Then run HTML validation and deploy as usual.
