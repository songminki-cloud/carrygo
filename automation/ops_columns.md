# CarryGo 운영용 시트 컬럼 보강

## 핵심 판단

신청폼에서 받는 값은 희망시간이고, 운영자가 실제 가능한 시간으로 묶어 최종 안내해야 한다.
따라서 Google Sheet에는 희망시간과 확정시간을 분리해야 한다.

## 추가/변경 컬럼

| 컬럼 | 용도 |
|---|---|
| pickup_preference | 고객이 선택한 희망 픽업 시간 |
| return_preference | 고객이 선택한 희망 반환 시간 |
| confirmed_pickup_time | 운영자가 확정한 픽업 시간 |
| confirmed_return_time | 운영자가 확정한 반환 시간 |
| pickup_location | 확정 픽업 장소 |
| return_location | 확정 반환 장소 |
| time_offer_sent_at | 가능 시간 안내 발송 시각 |
| payment_request_sent_at | 결제 요청 발송 시각 |
| payment_due_at | 결제 마감 시각. 기본: 안내 후 2시간 |
| invoice_id | PayPal invoice id 또는 수동 관리번호 |
| paid_at | 결제 확인 시각 |
| paid_amount | 결제 확인 금액 |
| payment_status | UNPAID / PAID / REFUNDED |
| confirmation_sent_at | 예약 확정 메시지 발송 시각 |

## 운영 흐름

1. `REQUESTED`: 고객 신청 접수
2. 운영자가 `confirmed_pickup_time`, `confirmed_return_time` 입력
3. `TIME_OFFERED`: 가능 시간 안내 완료
4. `payment_request_sent_at`, `payment_due_at`, `payment_link/invoice_id` 입력
5. `INVOICE_SENT`: 결제 요청 완료
6. 결제 확인 후 `PAID`
7. 확정 메시지 발송 후 `CONFIRMED`

## 고객 발송 메시지에 필요한 필드

- booking_no
- event_date
- bag_count
- confirmed_pickup_time
- confirmed_return_time
- pickup_location
- return_location
- amount_due
- payment_method
- payment_link 또는 PayPal invoice link
- payment_due_at
