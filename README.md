# mist-ap-ssid-channel-report

For every AP in a Mist org, shows **which SSIDs it publishes on which channel**, and **which Device
Profile is applied to it**, in one (big) Excel report.

## How it works

For every site in the org (or only the sites that match a name filter):

1. Pulls the live AP stats (`/sites/{site_id}/stats/devices?type=ap`). This gives each AP's Device Profile,
   effective radio config, and per radio: channel, width, power, clients, and `num_wlans` (how many
   SSIDs the radio is really broadcasting)
2. Pulls the WLANs that reach the site (`/sites/{site_id}/wlans/derived`, covering org WLAN Templates and site WLANs)
3. Works out which of those WLANs reach each AP:
   - the WLAN must be enabled
   - if its WLAN Template has **filter by Device Profile** set, the AP's Device Profile must be in the Template's list
   - if the WLAN is set to **apply to specific APs**, the AP must be one of them
   - the WLAN goes out on the bands in its `bands` setting
4. Checks the worked-out SSID count per band against the live `num_wlans`, so you can trust the list
5. Checks the live channel and power against the AP's allowed channels and power range

A dual-band radio (e.g. `band_24` running as 5GHz) is reported under the band it is really using.

## The Excel report

`ap_ssid_channel_report_<org>_<site filter or ALL-SITES>_<date_time>UTC.xlsx`

| Sheet | What's on it |
|---|---|
| Summary | Run details, counts, how everything is worked out, colour key |
| AP SSIDs by Channel | One row per AP: site, name, MAC, model, serial, firmware, status, uptime, **Device Profile**, then for 2.4 / 5 / 6GHz: **channel**, width, power, antenna gain, clients, **SSIDs**, SSID count (config) and SSID count (live), then the checks |
| AP SSID Detail | One row per AP / band / SSID: channel, width, power, whether the radio is on, SSID, WLAN ID, the WLAN Template (or site WLAN) it comes from, auth type, VLAN, hidden |
| Needs a Look | Connected APs where the SSID check failed, or the channel / power is outside the allowed limits |
| Device Profiles | Each Device Profile in use: AP count, per-band radio on/disabled, allowed channels, power range, width, and the WLAN Templates filtered to it |

### Reading the checks

- **SSID check (config vs live):** `OK` means the worked-out SSID count matches what the radio is really broadcasting.
  Amber means it doesn't, so the list may be wrong for that AP
- **Outside Device Profile limits** (red): the live channel isn't in the allowed list, or the power is above `power_max`
- **Power below Device Profile min (info):** not a fault. Mist lowers the power below `power_min` to stay
  within the legal EIRP limit, e.g. with a high-gain directional antenna or on a low-power channel
- `(radio disabled)` means the band is switched off in the Device Profile or AP config; `(radio off)` means
  the radio exists but isn't transmitting
- WLANs targeted with **WxLAN tags** can't be resolved per AP. They are listed with "(wxtag)" and those
  APs are not SSID-checked

## Setup

```
pip3 install -r requirements.txt
cp mist_ap_ssid_channel_report.ini.example mist_ap_ssid_channel_report.ini
# edit mist_ap_ssid_channel_report.ini: add api_token, org_id (and api_host if not EU)
python3 mist_ap_ssid_channel_report.py
```

Use a different settings file: `python3 mist_ap_ssid_channel_report.py other_org.ini`

`mist_ap_ssid_channel_report.ini` and `*.xlsx` are git-ignored, so tokens and results never get committed.

### Company proxies (Zscaler etc.)

If you see `certificate verify failed: unable to get local issuer certificate`, a company proxy is
probably inspecting HTTPS traffic. Settings go under `[network]` in the `.ini`:

- `use_system_certs = true` (default): trusts the Windows/macOS certificate store via `truststore`
  (installed by `requirements.txt`)
- `ca_bundle`: path to your company's root certificate (`.pem`)
- `http_proxy` / `https_proxy`: only if the `HTTP_PROXY` / `HTTPS_PROXY` environment variables aren't set
- `verify_ssl = false`: last resort only, turns certificate checking off

Leaving the `[network]` section out works as normal.

## API usage

About 2 calls per site plus 5 for the org (sites, WLAN Templates, Device Profiles). A 62-site,
3,600-AP org takes about 115 calls. HTTP 429 waits 60s and retries; 5xx and connection errors retry
3 times. Any other error **stops the script without writing a report**, so you never get a report
with sites quietly missing.
