# logs/

A timestamped log file is created here on every pipeline run.

```
YYYY-MM-DD_HH-MM-SS.log
```

Each file captures the full execution trace for that run:
`INFO` level and above for normal operation, `DEBUG` for detailed tracing.
Console output mirrors `INFO` level only.

Old log files are never deleted automatically — clean up manually as needed.

> This folder is git-ignored. Never commit log files to the repository.