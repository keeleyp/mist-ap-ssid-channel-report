# Changelog

## v1.0.0

First release.

- One row per AP: Device Profile, and per band (2.4 / 5 / 6GHz) the live channel, width, power,
  antenna gain, clients and the SSIDs published
- One row per AP / band / SSID with the WLAN ID, the WLAN Template it comes from, auth type, VLAN and hidden
- Works out SSIDs per AP from the site's derived WLANs, WLAN Template Device Profile filters and
  "apply to specific APs"; checks the result against the radio's live `num_wlans`
- Flags channels outside the allowed list and power above `power_max`; notes power below `power_min`
  (EIRP cap) separately
- Device Profiles sheet, Needs a Look sheet, Summary sheet
- Optional `[network]` section for company HTTPS proxies (Zscaler etc.)
- Stops without writing a report if any API call fails, rather than writing one with sites missing
