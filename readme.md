# Juniper Mist - EVPN Multihoming Fabric Builder

**Version:** 0.1  
**Author:** Lukas Eisenberger (leisenberger@juniper.net)

> ⚠️ **USE AT YOUR OWN RISK!** ⚠️  
> This tool is in early development (v0.1). Please test thoroughly in a non-production environment before deploying.

## Overview

EVPN Multihoming Fabric Builder automates the creation and update of EVPN-based multihoming topologies for **Juniper Mist**.  
It reads input data from Excel worksheets, builds or updates the corresponding topologies, and generates per-device configurations — ensuring a consistent, repeatable, and API-driven fabric deployment experience.

This version focuses on **EVPN multihoming** for collapsed-core or distribution designs, where dual-attached access switches connect redundantly to two or more cores.

## Features

### Core Capabilities
- **Excel-based Configuration Input**
  - Reads configuration data from Excel sheets: **FABRIC**, **INTERFACES**, and **NETWORKS**
  - Simplifies large-scale configuration management through spreadsheet-driven definitions
  - Detects and **updates existing topologies** instead of creating duplicates

### Device & Topology Management
- **Hostname-based Resolution**
  - Uses hostnames instead of device IDs for topology mapping  
  - Automatically resolves devices via the Mist API

- **Topology Detection & Update Mode**
  - Recognizes existing topologies by name  
  - Updates existing topologies and **preserves user-configured ports** during fabric updates

- **Smart Configuration Merging**
  - Caches device configurations before updates to intelligently merge port settings and avoid overwrites

### Network Configuration
- **Automatic Gateway and IP Assignment**
  - Corrects site “networks” settings, including IPv4/IPv6 gateways  
  - Assigns `other_ip_configs` dynamically:
    - For the first two core switches, IPs are automatically derived as `gateway+1` and `gateway+2`

- **Static Route Support**
  - Supports both IPv4 and IPv6 static routes  
  - Uses the format: `route@nexthop`
  - Allows multiple routes separated by spaces

### Interface Configuration
- **Port Speed and Channelization**
  - Reads speed and channelization settings directly from the **INTERFACES** sheet  
  - Supports:
    - Speeds: 10G, 25G, 50G, 100G, 200G, AUTO  
    - Channelization: TRUE/FALSE (for breakout ports)

## Prerequisites
- **Python:** 3.14 tested  
- **Dependencies:** `requests`, `openpyxl`

## Installation

```bash
# Clone the repository
git clone https://github.com/eisenluk/mist-campus-fabric-evpn-mh.git

# Navigate to the directory
cd mist-campus-fabric-evpn-mh

# Install dependencies
# pip install -r requirements.txt

```
## Usage

(Add usage instructions here, for example:)

0. Change values in excel spread sheet
1. Create an API token on your Mist Org
2. Copy API token and ORG-ID to the script
3. Run the script.

## Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

## License

(Add your license information here, e.g., MIT, Apache 2.0, etc.)

## Support

For questions, issues, or suggestions, please contact:
- **Author:** Lukas Eisenberger
- **Email:** leisenberger@juniper.net

## Disclaimer

This tool is provided as-is without any warranties. Always validate generated configurations before deploying to production environments. The author and contributors are not responsible for any network outages or issues resulting from the use of this tool.

---

*Last updated: [Date]*  
*Version: 0.1*
