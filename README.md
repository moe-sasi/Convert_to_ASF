# Convert_to_ASF

## Setup

Install dependencies before running the Streamlit app:

```
pip install -r requirements.txt
```

## Mapping overrides

The app ships with a default `mapping_overrides.yaml` file in the repository
root that is loaded automatically to guide fuzzy matching. To adjust alias
preferences, update that YAML file before running the app. Example additions:

```
LOAN_TYPE_LS:
  - Product Name
  - Product Type
BORROWER_NAME: Customer Name
```

Changes to `mapping_overrides.yaml` will be reflected the next time you start
the app.

## Constant field values

To hard-code specific ASF fields to a single value for every row, add them to
`constant_values.yaml` in the repository root:

```
# constant_values.yaml
SERVICING FEE %: 0.0025
MASTER SERVICER: Redwood
```

If present, these values are written even when a source column is not mapped.
