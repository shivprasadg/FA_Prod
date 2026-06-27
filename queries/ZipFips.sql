UPDATE
  Zipcty1
SET
  Zipcty1.State = Mid([Field1], 24, 2),
  Zipcty1.FIPS = Mid([Field1], 26, 3),
  Zipcty1.County = Mid([Field1], 29, 25),
  Zipcty1.Plus4 = Mid([Field1], 16, 4);
