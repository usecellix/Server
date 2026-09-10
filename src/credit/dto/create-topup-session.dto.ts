import { IsIn } from 'class-validator';

/** Body for `POST /billing/checkout/topup`. */
export class CreateTopupSessionDto {
  @IsIn(['small', 'medium', 'large'])
  packId!: 'small' | 'medium' | 'large';
}
