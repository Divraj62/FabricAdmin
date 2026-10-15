# Refresh Health Schema

The `PowerBI_RefreshHealth.csv` export contains one row per public-workspace semantic model assessed by Refresh Monitoring. The classifications and thresholds are produced by `Get-RefreshSequenceClassification` and `Get-RefreshHealthInventory`; this document describes the current output and does not redefine that logic.

| Column | Meaning |
| --- | --- |
| `Workspace`, `WorkspaceId` | Workspace name and identifier. |
| `SemanticModel`, `SemanticModelId` | Semantic model name and identifier. |
| `HealthColor`, `Health`, `CurrentRiskLevel` | Current status, display color, and evidence-based risk label. |
| `CurrentClassification`, `ClassificationReason` | Newest-refresh classification and its explanation. |
| `EvaluatedSequence`, `ConsecutiveFailures` | Evaluated newest-to-oldest status sequence and consecutive failure count. |
| `LatestStatus`, `IsRunning` | Latest refresh status and whether it is currently running. |
| `LastSuccessfulRefresh`, `LastFailedRefresh`, `FailureCount` | Recent success/failure timestamps and failures in the current evidence window. |
| `StartTime`, `EndTime`, `DurationMinutes`, `BaselineDurationMinutes` | Latest refresh timing and available historical duration baseline. |
| `RefreshType` | Power BI refresh type returned for the latest record. |
| `Server`, `Database`, `DatasourceType`, `Gateway` | Resolved datasource and gateway summary. |
| `DatasourceMappings` | JSON-encoded datasource connection metadata returned by Power BI. |
| `ScheduleHealth`, `ScheduleEnabled`, `ScheduleDays`, `ScheduleTimes`, `ScheduleTimeZone`, `ScheduleNotifyOption` | Refresh schedule metadata and assessment. |
| `MonitoringGaps` | Missing or unavailable metadata that limits assessment. |
| `ErrorCode`, `ErrorDescription`, `FailureCategory`, `Reason`, `ServiceExceptionJson`, `RequestId` | Latest refresh error details when available. |
| `RefreshAttempts`, `RecentFailureDetails`, `FailureHistoryDetails` | JSON-encoded attempt and failure evidence. |
| `CredentialAssessment` | Explicit limitation: credentials are not actively tested. |
| `AffectedReports` | Reports linked to the semantic model in the retrieved inventory. |
| `Evidence`, `RecommendedAction`, `Validation` | Supporting observations, administrator guidance, and validation note. |
| `SnapshotTimeUtc` | UTC time when the health snapshot was generated. |

## Current Classifications

- `HEALTHY`
- `RECOVERED`
- `INTERMITTENT / WARNING`
- `CURRENTLY RUNNING`
- `CURRENTLY FAILED / HIGH RISK`
- `UNKNOWN / NEEDS REVIEW`

Risk values currently include `LOW`, `WARNING`, `HIGH RISK`, and `UNKNOWN`. These are evidence-based indicators, not predictions of certain future failure. Missing metadata can make a result `UNKNOWN`; it does not prove datasource failure.