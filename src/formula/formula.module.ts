import { Module } from '@nestjs/common';
import { FormulaAnalyzer } from './formula.analyzer';
import { FormulaValidatorService } from './formula-validator.service';

@Module({
  providers: [FormulaAnalyzer, FormulaValidatorService],
  exports: [FormulaAnalyzer, FormulaValidatorService],
})
export class FormulaModule {}
