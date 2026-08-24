CAPACITY_QUERY = """
query CapacityReport {
  clusterConnection(first: 100) {
    nodes {
      name
      status
      metric {
        totalCapacity
        usedCapacity
        availableCapacity
        averageDailyGrowth
        lastUpdateTime
      }
    }
  }
}
"""
