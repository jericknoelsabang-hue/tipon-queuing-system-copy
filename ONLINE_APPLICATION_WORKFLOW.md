# Online Recruitment Application Process

This process keeps online applicants in a separate queue until they confirm arrival, so online appointments do not affect walk-in queue order.

```mermaid
flowchart TD
  A["Applicant opens Get Queue Page"] --> B["Chooses New Hire or Rehire"]
  B --> C["Completes online recruitment form and uploads pre-requirements"]
  C --> D["Submits application"]
  D --> E["System creates a private application link and stores the application"]
  E --> F["Recruiter reviews the form and uploaded documents in Online Queue"]
  F --> G{"Are the details and requirements complete?"}
  G -- "No" --> H["Recruiter messages applicant"]
  H --> I["Applicant replies in the private chat and provides missing details or files"]
  I --> F
  G -- "Yes" --> J["Recruiter marks the application complete"]
  J --> K["Applicant selects an available office date and time"]
  K --> L["Applicant gets an online queue number, such as OLNH-000001 or OLRH-000002"]
  L --> M["Applicant arrives on the appointment date"]
  M --> N["Applicant taps Nandito na ako sa office or confirms arrival by office QR"]
  N --> O["System moves the checked-in case into the Online Queue list"]
  O --> P["Recruiter calls the applicant to an assigned service window"]
  P --> Q["Recruiter opens Tipon Form for NH, RH, or TR"]
  Q --> R["Recruiter confirms details, captures photo and signature, then saves"]
  R --> S["Application and capture links are saved to ONLINE_APPLICATIONS"]
  S --> T["Recruiter completes the transaction"]
```

## Queue handling

- The online queue number is issued after the recruiter approves the requirements and the applicant books an appointment.
- Online applicants remain separate from walk-in applicants until they confirm arrival.
- Online queue numbering continues as applicants book; the prefix identifies the transaction (for example, `OLNH` or `OLRH`).
- Recruiter and applicant messages stay in the private application conversation while the application is being reviewed.
- The Tipon Form action is available only for New Hire, Rehire, and Transfer transactions.
