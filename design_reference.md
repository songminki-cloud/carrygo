# CarryGo Website Design Reference

Source package: `/Users/Jay/Downloads/carrygo_mobile_web_package`

Key files:
- `css/tokens.css`
- `css/styles.css`
- `assets/logo/carrygo_logo_final.png`
- `screens/home.html`
- `screens/reserve.html`
- `screens/confirmation.html`
- `preview/CARRYGO_mobile_web_design_board.png`

## Core visual language

CarryGo should feel like a **concert luggage claim tag system**, not a generic startup landing page.

Keywords:
- concert operations
- luggage claim tag
- ticket stub
- numbered tag
- event desk
- high-visibility field UI

## Palette

Use the package tokens:

```css
--cg-black: #050505;
--cg-ink: #111111;
--cg-charcoal: #262626;
--cg-paper: #ffffff;
--cg-warm-paper: #f7f4ef;
--cg-line: #d8d1c7;
--cg-muted: #736f68;
--cg-muted-red: #8f3c32;
```

Rules:
- Black/white/off-white first.
- Use warm paper background.
- No mint-heavy startup UI.
- Minimal accent color only when operationally useful.

## Typography

- Bold grotesk/sans-serif feel.
- Use Inter / Pretendard / Apple SD Gothic Neo / Noto Sans KR stack.
- Headlines: huge, heavy, tight line height.
- Labels: small uppercase, wide tracking.
- Claim numbers / reservation IDs: oversized, bold.

Example hierarchy:

```text
Leave it.
Enjoy the show.
Pick it up.
```

Labels:

```text
CONCERT LUGGAGE CARE · SEOUL
TAG NO.
BASE FEE
PICKUP & DROP
NEXT-DAY PICKUP
```

## Layout

- Mobile-first around 390px base width.
- Single-column stacked layout.
- Generous vertical rhythm.
- Thick black outlines.
- Rounded cards.
- Ticket/stub/perforation details.
- Avoid wide desktop-style marketing sections as the primary design basis.

Homepage can expand to desktop, but the master design should be mobile.

## Components to reuse/adapt

From the package:

- `.logo-lockup`
- `.ticket-card`
- `.ticket-head`
- `.ticket-body`
- `.ticket-row`
- `.ticket-mini`
- `.stub`
- `.claim-number`
- `.tag-label`
- `.status-pill`
- `.step-grid`
- `.form-card`
- `.field`
- `.input`
- `.btn.primary`
- `.btn.secondary`
- `.qr-box`

## Public landing page structure

Use the mobile package style, but adapt content to current MVP operations.

Recommended sections:

1. Topbar
   - Real logo image
   - Minimal menu or CTA

2. Hero
   - `Leave it. Enjoy the show. Pick it up.`
   - Korean support copy
   - CTA: Reserve Bag Drop

3. Active concert ticket card
   - SHINee WORLD VII
   - KSPO DOME
   - dates

4. Base Fee ticket/card
   - `₩20,000 / $15`
   - Includes: 1 suitcase, same-day pickup & drop, within 2 hours after concert, max 28 inch / 23kg

5. How it works
   - Apply
   - Pay Base Fee
   - Receive QR
   - Show QR onsite
   - Pick up after concert

6. Rules / field operations
   - Extra bag cash only
   - Next-day pickup
   - Refund policy
   - Luggage Tag matching

7. Reservation form
   - Connected to Apps Script `create_reservation`
   - Use package form-card style

8. FAQ
   - Short operational answers only

## Content rules

Current MVP facts to reflect:

- Base Fee: `₩20,000 / $15`
- Includes: 1 suitcase, same-day pickup & drop, pickup within 2 hours after concert, max 28 inch / 23kg
- Extra bag/shopping bag: `₩10,000`, onsite cash only. If no KRW, `$10` cash. No exchange-rate calculation.
- Next-day Pickup: if not picked up within 2 hours after concert. Next day 10:00–12:00 at CarryGo-designated place. Fee same as Base Fee, once per reservation.
- Refund: no refund for customer cancellation/no-show after payment. Full refund if concert cancelled or CarryGo cannot provide service.
- Payment methods: KakaoPay, PayPal, optional Bank Transfer.
- Public page should not expose detailed account/private payment info.
- QR is sent after payment confirmation.
- Luggage Tag is physical field matching tool.

## Avoid

- Old mint gradient startup style.
- Soft lifestyle copy.
- Photos as main hero.
- Thin typography.
- Complex navigation.
- Too much legal text above the fold.
- Multilingual tabs before Korean/English core is stable.
- Placeholder/fake links that look final.
- Public detailed payment account info.

## Design judgment

The current `index.html` should be rebuilt using this package, not patched from the old landing page. The old page was based on outdated pricing/operations and has the wrong visual language.


## Border / outline rules

CarryGo uses borders as an operational structure, not decoration.

### Use strong borders only for

- Major work areas: page shell, action card, reservation list card, modal card.
- Individual operational records: reservation rows, concert rows, QR/check-in result cards.
- Primary interactive controls that must read as buttons.

These borders mean: **this is a unit of work or a thing the staff can act on.**

### Do not use borders for

- Status text such as `관리자 모드 활성화`.
- Count/status messages such as `예약취소 대상 2건`.
- Warning/helper copy.
- Section labels such as `Menu`.
- The `Admin` wordmark beside the logo.

These should be handled with text weight, color, position, or small status dots.

### Color roles

- Black / ink: structure, primary actions, selected state.
- Muted red: important operational warnings and count/status text that requires attention. No box by default.
- Green: only a small status dot for active/online/unlocked state. No green cards or green borders.
- Muted gray: secondary explanations.

### Admin-specific rule

The admin screen is a field-operation tool, not a landing page. Keep the header compact. Separate:

1. identity/menu area
2. action area
3. record/list area

Do not wrap every line in a pill or card. If everything has a border, nothing is important.
