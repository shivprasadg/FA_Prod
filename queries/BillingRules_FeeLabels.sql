SELECT
  BillingRules.RuleType,
  BillingRules.RuleName,
  BillingRules.RuleString,
  BillingRules.RuleValue
FROM
  BillingRules
WHERE
  (
    (
      (BillingRules.RuleType)= "FeeLabel"
    )
  );
