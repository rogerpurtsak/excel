"""Test fixtures.

Data extracted from samples/tooleht5.xls:
  Column A "Automaatrelvad" (registered automatic firearms, thousands per 100k)
  Column B "Mõrvad"         (homicides per 100k)
These 8 pairs reproduce the reference statistics in the task spec.
"""

# x = firearms, y = homicides
FIREARMS = [11.6, 8.3, 3.6, 0.6, 6.9, 2.5, 2.4, 2.6]
HOMICIDES = [13.1, 10.6, 10.1, 4.4, 11.5, 6.6, 3.6, 5.3]

# Reference values the math engine must reproduce (tolerance 1e-6).
REF_PEARSON_R = 0.88498417
REF_INTERCEPT = 4.04725078
REF_SLOPE = 0.85251932
REF_R_SQUARED = 0.78319698
REF_PRED_AT_12_5 = 14.70374226
