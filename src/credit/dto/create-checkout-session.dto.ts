import { IsIn } from 'class-validator';

/** Body for `POST /billing/checkout/subscribe`. */
export class CreateCheckoutSessionDto {
  @IsIn(['solo', 'firm'])
  planTier!: 'solo' | 'firm';
}
