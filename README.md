[![License](https://img.shields.io/github/license/toolarium/toolarium-outlook-exporter)](https://github.com/toolarium/toolarium-outlook-exporter/blob/main/LICENSE)


# toolarium-outlook-exporter

It gives you access to your Outlook calendar and emails via an export. It does the job that you won't find at M$.

## Getting Started

Download the package and use the batch file **bin\outlook-exporter**. There is a help, see **bin\outlook-exporter --help**

### Samples:

- Export all calendar entries and sent mails from current month: ```cmd bin\outlook-exporter```
- Export all calendar entries and sent mails from july of the current year: ```cmd bin\outlook-exporter 7```
- Export all calendar entries and sent mails from july from the year 2024: ```cmd bin\outlook-exporter 7.2024```
- Export all calendar entries and sent mails from 2nd of July only: ```cmd bin\outlook-exporter 2.7.2025```
- Export all calendar entries and sent mails from 2nd of July until 4th of July: ```cmd bin\outlook-exporter 2.7.2025 4.7.2025```


## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/toolarium/toolarium-changelog-parser/tags).