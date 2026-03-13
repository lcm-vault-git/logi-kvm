# logi-kvm
Python script to synchronize Logitech Unifying devices channel switch and monitor input change.

Provides switching of Logitech Unifying devices channels and VCP (VESA MCCS) compliant monitors input selection as a reaction to Logitech Easy Switch keys events.

Notes:
- Unifying channels indexes starts at 0
- Unifying receiver slots indexes starts at 1
- Display number and input name has to be found by try and error.

**Works on Windows 10
Unfortunately hidapi is not working properly on Linux yet**

Any Unifying Hardware should be compatible. In order to add other Unifying devices support please read this whole readme file.

Default configuration is created in function populate_config.

## Usage:
```
python switch-input.py -h
python switch-input.py <CHANNEL> -c config.json
```

- `CHANNEL`: this PC's Unifying channel number (0-indexed)
  - PC 1 → `0`
  - PC 2 → `1`

Example:
```bash
# PC 1
python switch-input.py 0 -c config.json

# PC 2
python switch-input.py 1 -c config.json
```

The config file (`config.json`) is shared across both PCs. Only the channel argument differs.

### Windows startup (background)
Create a `.vbs` file in `shell:startup`:
```vbs
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d C:\path\to\logi-kvm && python switch-input.py <CHANNEL> -c config.json", 0, False
```

## Ping-based keyboard detection (for keyboards without Easy-Switch events)

Some keyboards (e.g. ERGO K860) do **not** send Easy-Switch key events via HID++. For these keyboards, the script uses a ping-based detection method:

1. Periodically sends HID++ 2.0 ping to the keyboard slot
2. If the keyboard responds → it's still connected to this PC
3. If ping misses exceed threshold → keyboard has switched to another PC → mouse is automatically switched to follow

This requires the script to run on **both PCs**:
- When keyboard leaves PC 2 → PC 2 detects ping miss → sends mouse to PC 1
- When keyboard leaves PC 1 → PC 1 detects ping miss → sends mouse to PC 2

**Battery note**: Constant pinging prevents the keyboard from entering deep sleep. The default ping interval (0.3s) prioritizes fast switching over battery life. Increase `PING_INTERVAL` in the script for better battery life (e.g. 2-3s).

## Adding other Unifying devices with change host capability
Best tool to discover how to make unifying device change host is to use [Solaar](https://github.com/pwr-Solaar/Solaar) on linux and list all field of devices by calling:

`solaar show`

Then you need to look for feature called `CHANGE HOST`.
This number (e.g. for MX Keys it's `9` and for MX Ergo it's `21`) has to be byte number 2 in `switch_message`. Byte 3 is for my current knowledge a magic number and has to be found by trial and error.
For detecting if device is sending any messages when easy switch key is pressed the easiest way is to use [hidapitester](https://github.com/todbot/hidapitester) to listen to the device. For MX Keys best results are for:

`hidapitester --vidpid 046D:C52B --usagePage 0xFF00 --usage 0x0002 -l 11 -t 5000 --open --read-input-forever`

Output for MX Keys is:
```
Opening device, vid/pid:0x046D/0xC52B, usagePage/usage: FF00/2
Device opened
Reading 11-byte input report 0, 5000 msec timeout...read 11 bytes:
 11 01 08 20 00 D2 01 00 00 00 00
```
Format of this message is:
| Header | Slot | Const | Const | Const | Key  | Key state |
|--------|------|-------|-------|-------|------|-----------|
|  0x11  | 0x01 | 0x08  | 0x20  | 0x00  | 0xD2 |  0x01     |

This script assumes that kay code is in byte 5 (Key) and therefore value in config doesn't matter. Received byte will be looked in `easy_switch_keys` list and if found, index will be taken as target channel number

Format of `switch_message` as for MX Ergo example is:
| Header | Slot | CHANGE HOST | Magic number | Channel | Padding | Padding |
|--------|------|-------------|--------------|---------|---------|---------|
|  0x10  | 0x02 |     0x15    |     0x1b     |   0x01  |  0x00   |  0x00   |

Byte 4 (Channel) will be replaces with channel number detected by `switch_detect_message`

`easy_switch_keys` are values of easy switch keys as received in the message, for MX Keys those are:
| Key | Value |
|-----|-------|
|  1  | 0xD1  |
|  2  | 0xD2  |
|  3  | oxD3  |

## Configuring monitors:
Monitors are controlled by [**VESA Monitor Control Command Set Standard**](https://milek7.pl/ddcbacklight/mccs.pdf).

Nice tool to play around and make sure that VCP message number and values for input change are correct is [ControlMyMonitor](https://www.nirsoft.net/utils/control_my_monitor.html)

## Example config files:

**IMPORTANT: JSON is sensitive for commas at the last field of object!**

**No monitor, ERGO K860 and MX Anywhere 2S (ping-based, 2 PCs):**
```json
{
    "monitors": [],
    "unifying_devices": [
        {
            "slot_id": 1,
            "dev_type": "ERGO K860",
            "switch_detect_message": [
                17, 1, 8, 32, 0, 255, 1, 0, 0, 0, 0
            ],
            "easy_switch_keys": [
                209, 210, 211
            ],
            "switch_message": [
                16, 1, 9, 30, 255, 0, 0
            ],
            "max_channels": 2
        },
        {
            "slot_id": 2,
            "dev_type": "MX Anywhere 2S",
            "switch_detect_message": [],
            "easy_switch_keys": [],
            "switch_message": [
                16, 2, 8, 30, 255, 0, 0
            ],
            "max_channels": 2
        }
    ],
    "unifying_channel": 0
}
```
Run with `switch-input.py 0` on PC 1, `switch-input.py 1` on PC 2.

**One monitor, MX keys and MX Ergo:**
```json
{
    "monitors": [
        {
            "channel_to_monitor_id": {
                "0": 2,
                "1": 1
            },
            "vcp_message_number": 60,
            "channel_to_input_dict": {
                "0": "HDMI2",
                "1": "HDMI1"
            }
        }
    ],
    "unifying_devices": [
        {
            "slot_id": 1,
            "dev_type": "MX Keys"
        },
        {
            "slot_id": 2,
            "dev_type": "MX Ergo"
        }
    ],
    "unifying_channel": 0
}
```

**Two monitors, MX keys and MX Ergo:**

Note that here `switch_detect_message` and `easy_switch_keys` are empty, as MX Ergo doesn't provide feedback on channel change
```json
{
    "monitors": [
        {
            "channel_to_monitor_id": {
                "0": 2,
                "1": 1,
                "2": 1,
            },
            "vcp_message_number": 60,
            "channel_to_input_dict": {
                "0": "DVI2",
                "1": "HDMI2",
                "2": "HDMI1"
            }
        }
        {
            "channel_to_monitor_id": {
                "0": 1,
                "1": 1,
                "2": 1
            },
            "vcp_message_number": 60,
            "channel_to_input_dict": {
                "0": "HDMI2",
                "1": "HDMI1",
                "2": "DVI1"
            }
        }
    ],
    "unifying_devices": [
        {
            "slot_id": 1,
            "dev_type": "MX Keys"
        },
        {
            "slot_id": 2,
            "dev_type": "MX Ergo",
            "switch_detect_message": [],
            "easy_switch_keys": [],
            "switch_message": [
                16,
                2,
                21,
                27,
                255,
                0,
                0
            ],
            "max_channels": 2
        }
    ],
    "unifying_channel": 0
}
```

**Two monitors, MX Master 3 and some other mx device, not kardcoded in the script:**
```json
{
    "monitors": [
        {
            "channel_to_monitor_id": {
                "0": 2,
                "1": 1,
                "2": 1,
            },
            "vcp_message_number": 60,
            "channel_to_input_dict": {
                "0": "DVI2",
                "1": "HDMI2",
                "2": "HDMI1"
            }
        }
        {
            "channel_to_monitor_id": {
                "0": 1,
                "1": 1,
                "2": 1
            },
            "vcp_message_number": 60,
            "channel_to_input_dict": {
                "0": "HDMI2",
                "1": "HDMI1",
                "2": "DVI1"
            }
        }
    ],
    "unifying_devices": [
        {
            "slot_id": 1,
            "dev_type": "MX Master 3"
        },
            "slot_id": 1,
            "dev_type": "Some other MX device",
            "switch_detect_message": [
                17,
                1,
                8,
                32,
                0,
                255,
                1
            ],
            "easy_switch_keys": [
                209,
                210,
                211
            ],
            "switch_message": [
                16,
                1,
                9,
                30,
                255,
                0,
                0
            ],
            "max_channels": 3
    ],
    "unifying_channel": 0
}
```
