# Concept models served by the app

Put domain models here, one folder per domain, for example:

```
concept-models/
└── apl/
    ├── business.yaml
    └── processes.yaml
```

Every `*.yaml` / `*.yml` file below this folder is served by `/api/concept-models`
and loaded automatically in the **Concepts** and **Processes** views when the
viewer has nothing loaded yet. Set the `CONCEPT_MODELS_DIR` environment variable
in `app.yaml` to read them from another folder inside the deployed app source
(for example one your deployment copies in from your models repository). The app
can only read files that are deployed with it.

This repository is public: do not commit organisation-internal models here. Copy
them in at deploy time from a private repository instead, or load them in the
browser with **Load YAML…**.

The YAML is read-only here; edit it in its own repository.
