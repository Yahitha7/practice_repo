import pandas as pd
import requests
import urllib3
from datetime import datetime

from config import GRAPHQL_URL
from graphql_queries import CAPACITY_QUERY
from cluster_mapping import CLUSTER_MAPPING, TH_HIGH, TH_LOW


urllib3.disable_warnings(
    urllib3.exceptions.InsecureRequestWarning
)


# ============================================================
# BYTES TO GB
# ============================================================

def bytes_to_gb(value):

    if value is None:
        return 0.0

    try:
        return float(value) / (1024 ** 3)

    except (TypeError, ValueError):
        return 0.0


# ============================================================
# GET CAPACITY FROM RUBRIK
# ============================================================

def get_capacity(token):

    headers = {
        "Authorization": "Bearer {}".format(token),
        "Content-Type": "application/json"
    }

    payload = {
        "query": CAPACITY_QUERY
    }

    print()
    print("Connecting to Rubrik GraphQL...")

    response = requests.post(
        GRAPHQL_URL,
        headers=headers,
        json=payload,
        verify=False,
        timeout=120
    )

    response.raise_for_status()

    result = response.json()

    if "errors" in result:

        raise RuntimeError(
            "Rubrik GraphQL error: "
            + str(result["errors"])
        )

    return (
        result
        .get("data", {})
        .get("clusterConnection", {})
        .get("nodes", [])
    )


# ============================================================
# BUILD DATAFRAME
# ============================================================

def build_dataframe(data):

    report_date = datetime.now().strftime(
        "%m/%d/%Y"
    )

    rows = []

    # --------------------------------------------------------
    # Process only configured clusters
    # --------------------------------------------------------

    for cluster in data:

        cluster_name = cluster.get("name")

        if cluster_name not in CLUSTER_MAPPING:
            continue

        mapping = CLUSTER_MAPPING[cluster_name]

        # ----------------------------------------------------
        # Validate mapping
        # ----------------------------------------------------

        required = [
            "Location",
            "Type",
            "Target Storage",
            "Region",
            "Country",
            "City"
        ]

        missing = [
            field
            for field in required
            if field not in mapping
        ]

        if missing:

            raise ValueError(
                "Incomplete mapping for {}: {}".format(
                    cluster_name,
                    ", ".join(missing)
                )
            )

        # ----------------------------------------------------
        # Capacity
        # ----------------------------------------------------

        metric = cluster.get("metric") or {}

        total_capacity = bytes_to_gb(
            metric.get("totalCapacity")
        )

        used_capacity = bytes_to_gb(
            metric.get("usedCapacity")
        )

        free_capacity = bytes_to_gb(
            metric.get("availableCapacity")
        )

        if (
            free_capacity <= 0
            and total_capacity > used_capacity
        ):

            free_capacity = (
                total_capacity
                - used_capacity
            )

        free_capacity = max(
            0.0,
            free_capacity
        )

        # ----------------------------------------------------
        # Percentages
        # ----------------------------------------------------

        if total_capacity > 0:

            used_percentage = (
                used_capacity
                /
                total_capacity
            )

            free_percentage = (
                free_capacity
                /
                total_capacity
            )

        else:

            used_percentage = 0.0
            free_percentage = 0.0

        # ----------------------------------------------------
        # Active
        # ----------------------------------------------------

        active = (
            "Yes"
            if cluster.get("status") == "Connected"
            else "No"
        )

        # ----------------------------------------------------
        # Create row
        # ----------------------------------------------------

        rows.append({

            "Month": report_date,

            "Active": active,

            "Location": mapping["Location"],

            "Type": mapping["Type"],

            "Target Storage": mapping["Target Storage"],

            "Region": mapping["Region"],

            "Country": mapping["Country"],

            "City": mapping["City"],

            "Total Capacity": round(
                total_capacity,
                2
            ),

            "Used Capacity": round(
                used_capacity,
                2
            ),

            "Free Capacity": round(
                free_capacity,
                2
            ),

            "Used %": round(
                used_percentage,
                6
            ),

            "Free %": round(
                free_percentage,
                6
            ),

            "TH High": TH_HIGH,

            "TH Low": TH_LOW
        })

    # ========================================================
    # DATAFRAME
    # ========================================================

    df = pd.DataFrame(rows)

    if df.empty:

        raise ValueError(
            "No configured Rubrik clusters were returned."
        )

    # --------------------------------------------------------
    # Column order
    # --------------------------------------------------------

    columns = [
        "Month",
        "Active",
        "Location",
        "Type",
        "Target Storage",
        "Region",
        "Country",
        "City",
        "Total Capacity",
        "Used Capacity",
        "Free Capacity",
        "Used %",
        "Free %",
        "TH High",
        "TH Low"
    ]

    df = df[columns]

    # ========================================================
    # VALIDATION
    # ========================================================

    expected = len(
        CLUSTER_MAPPING
    )

    actual = len(df)

    print()
    print("=" * 70)
    print("Cluster Validation")
    print("=" * 70)

    print(
        "Clusters returned by Rubrik : "
        + str(len(data))
    )

    print(
        "Clusters configured          : "
        + str(expected)
    )

    print(
        "Clusters included in report  : "
        + str(actual)
    )

    if actual != expected:

        missing_clusters = (
            set(CLUSTER_MAPPING.keys())
            -
            set(df["Target Storage"])
        )

        print()
        print("Missing configured clusters:")

        for cluster in sorted(
            missing_clusters
        ):

            print(
                "  - "
                + cluster
            )

        raise ValueError(
            "Not all configured clusters were returned."
        )

    print(
        "All configured clusters are present."
    )

    print("=" * 70)

    # ========================================================
    # REGION VALIDATION
    # ========================================================

    print()
    print("=" * 70)
    print("Region Validation")
    print("=" * 70)

    print(
        df["Region"]
        .value_counts(
            dropna=False
        )
        .to_string()
    )

    print("=" * 70)

    # --------------------------------------------------------
    # Fail if Region is missing
    # --------------------------------------------------------

    if df["Region"].isna().any():

        raise ValueError(
            "Region mapping contains NaN values."
        )

    # ========================================================
    # SORT IN CLUSTER MAPPING ORDER
    # ========================================================

    cluster_order = {
        cluster: index
        for index, cluster in enumerate(
            CLUSTER_MAPPING.keys()
        )
    }

    df["_order"] = (
        df["Target Storage"]
        .map(cluster_order)
    )

    df = (
        df
        .sort_values("_order")
        .drop(columns=["_order"])
        .reset_index(drop=True)
    )

    return df
