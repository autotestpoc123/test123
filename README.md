flowchart TD
    A[Start validation] --> B[Resolve analysis summary<br/>AnalysisSummaryPath / RunId / DateRange / LatestPointer]
    B --> C[Load analysisSummary and taskSummaries]
    C --> D[Build relatedRuns by same StartDate + EndDate]
    D --> E{dynamicSummaryTaskNames.Count > 0<br/>and relatedRuns.Count > 0 ?}

    E -- Yes --> F[Get-MergedTaskSummaries<br/>from related runs]
    E -- No --> G[Use current run taskSummaries only]

    F --> H[Collect usedRunIds = unique RunId<br/>from mergedSummaryData.TaskSources]
    G --> H2[usedRunIds = empty]
    H2 --> I[Get selectedRunId<br/>analysisSummary.RunId else runFolder name]
    H --> I

    I --> J{usedRunIds.Count > 1 ?}
    J -- Yes --> K[validationMode = aggregated]
    J -- No --> L{usedRunIds.Count == 1<br/>and selectedRunId exists<br/>and usedRunIds[0] != selectedRunId ?}
    L -- Yes --> K
    L -- No --> M[validationMode = single-run]
    J -- No --> L

