let
    Staging = qCanonicalStaging,
    Rejected = Table.SelectRows(Staging, each [QualityStatus] = "FAIL")
in
    Rejected