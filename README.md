Email template engine for dynamics 365

# Example

Привет.
Недавно вы участвовали в интервью с несколькими кандидатами на вакансию %;ym_name;%. Пожалуйста, перейдите по ссылкам ниже и оставьте свой фидбэк.
%;RELATED_ym_history-ym_applicationid-ym_candidateid-RECORD_URL|Оставить фидбэк;%

&;comment;&

```
var tm = new Templator.TemplateManager(service, templateName, _secureConfig);
tm.AddFilter("ym_history-ym_applicationid", GetQuery(selectedCandidatesGuids, context.PrimaryEntityId));
var email = tm.FillEmailByTemplate(context.PrimaryEntityName, context.PrimaryEntityId, fillDict);
```

# About

Support:
* Placeholder for entity field

```
%;ym_name;%
```

* Placeholder for N:1 relationship

```
%;relatedrecord.ym_name;%
```

* Placeholder for 1:N 
* Links to record
