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


### Configuration:

- In the config folder (which will be created if it does not exist) are two files:
    - calendar-subject-filter.txt: It defines subject filter of entries you don't like to export. It acts as a keyword filter.
    - calendar-attendee-filter.txt: It defines line by line attendee to fliter out from the export. 

#### Configuration customer-filter

- In the subfolder “config\customer-filter”, you can define output-specific filters that you want to use to split entries as example for customers. This acts as a filter. Each word inside this file (line by line) defines keywords which are mapping to this specific output. If you have two different you simply create two different files with it's name: **customer-a.txt** amd **customer-b.txt**.
- In addition you can have additional attendee filter for each customer. It defines line by line attendee to fliter for this specific export, e.g. **customer-a-attendees.txt** amd **customer-b-attendees.txt**.
- In the output to each mail a specific time is applied (by default 0.5 hour). This you can customize for each split, e.g. **customer-a-duration.txt** amd **customer-b-duration.txt**. Inside this file you can configure the following keywords (key / value). The value corresponds to the spent hours:
    DEFAULT_EMAIL_DURATION=0.25
    ADDITIONAL_EMAIL_DURATION=0.15

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/toolarium/toolarium-outlook-exporter/tags).