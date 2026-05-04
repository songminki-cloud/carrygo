# CarryGo 결제확인 자동화 설계

## 결론

PayPal은 자동화 가능하다. 카카오페이 개인 송금은 공식 API가 없어 완전 자동화가 어렵다.
따라서 1차 MVP는 아래 구조로 간다.

1. 홈페이지 신청폼 → Google Sheet 자동 저장
2. 예약번호 자동 생성
3. 운영자가 가능한 시간 안내 + 결제 링크 발송
4. PayPal 결제 완료 메일 또는 Webhook → Sheet 자동 업데이트
5. 카카오페이/계좌 송금 → 고객이 예약번호 입력 + 운영자 빠른 확인
6. 결제 확인된 건만 `CONFIRMED`

## 권장 상태값

- `REQUESTED`: 신청 접수
- `TIME_OFFERED`: 가능 시간 안내 완료
- `INVOICE_SENT`: PayPal invoice 또는 KakaoPay 안내 발송
- `PAYMENT_REPORTED`: 고객이 송금/결제했다고 회신
- `PAID`: 결제 확인 완료
- `CONFIRMED`: 예약 확정 안내 발송 완료
- `WAITLIST`: 대기
- `CANCELLED`: 취소
- `NO_SHOW`: 노쇼

## Google Sheet 컬럼

| 컬럼 | 설명 |
|---|---|
| created_at | 신청 시각 |
| booking_no | 예약번호, 예: CG-0529-001 |
| status | 예약 상태 |
| event_date | 공연일 |
| name | 고객명 |
| email | 이메일 |
| phone | 휴대폰 |
| language | 언어 |
| bag_count | 짐 개수 |
| pickup_preference | 희망 픽업 시간 |
| return_preference | 희망 반환 시간 |
| flexibility | 시간 조정 가능 범위 |
| payment_method | KakaoPay / PayPal / Cash |
| currency | KRW / USD |
| amount_due | 결제 요청 금액 |
| payment_link | PayPal invoice link 또는 KakaoPay link |
| invoice_id | PayPal invoice id 또는 수동 관리번호 |
| paid_at | 결제 확인 시각 |
| paid_amount | 결제 확인 금액 |
| payment_status | UNPAID / PAID / REFUNDED |
| operator_note | 운영자 메모 |

## PayPal 자동화 후보

### 1차: Gmail 파싱
PayPal 결제 완료 이메일에서 예약번호(`CG-0529-001`) 또는 invoice id를 읽어 Sheet를 업데이트한다.
장점: 서버/Webhook 없이 Apps Script만으로 가능.
단점: PayPal 이메일 포맷 변경에 취약. 처음 며칠은 수동 검증 필요.

### 2차: PayPal Webhook
PayPal Developer 앱에서 `INVOICING.INVOICE.PAID` 이벤트를 Apps Script Web App으로 보내 Sheet 업데이트.
장점: 정석.
단점: PayPal Developer 설정과 Webhook 검증이 필요.

## KakaoPay 자동화 현실

개인 카카오페이 송금 링크/계좌 송금은 공식 수신 API가 없다.
완전 자동화는 어렵다.

가능한 방법:
1. 고객에게 송금 메모/입금자명에 예약번호 입력 강제
2. 고객이 결제 완료 버튼/폼 제출 → Sheet `PAYMENT_REPORTED`
3. 운영자가 카카오페이/은행 앱에서 확인 후 `PAID`

은행/카카오페이 알림이 이메일로 온다면 Gmail 파싱 자동화 가능하지만, 알림 출처/본문 포맷 확인 전까지는 확정 불가.

## 운영 원칙

- 결제 완료자만 예약 확정
- 결제 대기 시간은 2시간
- 결제 미확인 건은 자동 취소 가능
- 예약번호 없는 송금은 확정하지 않음
- PayPal invoice는 고객별로 새로 생성
