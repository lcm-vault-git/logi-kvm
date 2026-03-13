#!/usr/bin/env python
"""Provides switching Logitech Unifying devices channels and VCP compliant monitors input switch
as a reaction to Logitech Easy Switch keys events

Notes:
    Unifying channels indexes starts at 0
    Unifying receiver slots indexes starts at 1
    Any Unifying Hardware should be compatible. In order to check what message has to be sent to the device to
    make it switch channel user has to check number of feature called "CHANGE HOST". This can be checked using
    software Solaar by listing devices via command 'solaar show'. This number has to be byte number 2 in
    switch_message in UnifyingDevice class. Byte number 3 in this message is, as for authors current knowledge,
    a magic byte and has to be found by try and error.
    Display number and input name has to be found by try and error.

    Default configuration is created in function populate_config
"""

import argparse
import sys
import logging
from logging.handlers import RotatingFileHandler
import monitorcontrol
import hid
import json
import time

__author__ = "Konrad Foit"
__copyright__ = "Copyright 2022, Konrad Foit"
__credits__ = ["https://github.com/Logitech/logi_craft_sdk/issues/28"]

__license__ = "MIT"
__maintainer__ = "Konrad Foit"
__status__ = "Development"

UNIFYING_RECEIVER_VID = 0x046D
UNIFYING_RECEIVER_PID = 0xC52B
UNIFYING_RECEIVER_LISTEN_USAGE_PAGE = 0xFF00
UNIFYING_RECEIVER_SEND_USAGE_PAGE = 0xFF00
UNIFYING_RECEIVER_LISTEN_USAGE = 0x0002
UNIFYING_RECEIVER_SEND_USAGE = 0x0001

listen_device_path = ""
send_device_path = ""

# Persistent HID handles
listen_handle = None
send_handle = None


def open_hid_handles():
    global listen_handle, send_handle
    listen_handle = hid.device()
    listen_handle.open_path(listen_device_path)
    listen_handle.set_nonblocking(1)
    send_handle = hid.device()
    send_handle.open_path(send_device_path)
    send_handle.set_nonblocking(0)


def unifying_listen(timeout_ms=500):
    start = time.time()
    while (time.time() - start) * 1000 < timeout_ms:
        data = listen_handle.read(64)
        if data:
            return data
        time.sleep(0.01)
    return []


def unifying_write(data):
    send_handle.write(data)


def ping_device(slot_id):
    """Send HID++ 2.0 ping (IRoot feature 0x00, function 1) to check if device is present.
    Returns True if device responds, False otherwise.
    """
    try:
        # HID++ 2.0 ping: feature_index=0x00 (IRoot), function=1 (ping), sw_id=0
        # Response echoes back ping_data (0xAA) if device is present
        unifying_write([0x10, slot_id, 0x00, 0x10, 0x00, 0x00, 0xAA])
    except Exception:
        return False
    # Read responses for up to 200ms, looking for one from our slot
    start = time.time()
    while (time.time() - start) * 1000 < 200:
        data = listen_handle.read(64)
        if data and len(data) >= 2:
            logging.debug("Ping response: %s" % str(data))
            if data[1] == slot_id:
                return True
        if not data:
            time.sleep(0.01)
    return False


def usb_discover():
    for dev in hid.enumerate(UNIFYING_RECEIVER_VID, UNIFYING_RECEIVER_PID):
        if dev['usage'] == UNIFYING_RECEIVER_LISTEN_USAGE and dev['usage_page'] == UNIFYING_RECEIVER_LISTEN_USAGE_PAGE:
            global listen_device_path
            listen_device_path = dev['path']
        elif dev['usage'] == UNIFYING_RECEIVER_SEND_USAGE and dev['usage_page'] == UNIFYING_RECEIVER_SEND_USAGE_PAGE:
            global send_device_path
            send_device_path = dev['path']


class UnifyingDevice:
    def __init__(self, dev_type, slot_id, switch_detect_message, easy_switch_keys, switch_message, max_channels):
        self.dev_type = dev_type
        self.slot_id = slot_id
        self.switch_detect_message = switch_detect_message
        self.easy_switch_keys = easy_switch_keys
        self.switch_message = switch_message
        self.max_channels = max_channels

    @staticmethod
    def get_from_type(dev_type, slot_id):
        if dev_type.lower() == 'MX Keys'.lower():
            # byte 0 0x11 is header
            # byte 1 contains index of a slot on which device is paired in unifying receiver
            # byte 2 is always 0x08
            # byte 3 is always 0x20
            # byte 4 is always 0x00
            # byte 5 contains key number. For MX Keys easy switch keys are 0xD1, 0xD2, 0xD3
            # byte 6 is key state 0x01 is pressed, 0x00 is released
            switch_detect_message = [0x11, slot_id, 0x08, 0x20, 0x00, 0xFF, 0x01]
            easy_switch_keys = [0xD1, 0xD2, 0xD3]
            # byte 4 (0xFF) is a placeholder for target channel number
            switch_message = [0x10, slot_id, 0x09, 0x1e, 0xFF, 0x00, 0x00]
            max_channels = 3
        elif dev_type.lower() == 'MX Ergo'.lower():
            # MX Ergo doesn't provide event on channel switch button
            switch_detect_message = []
            easy_switch_keys = []
            # byte 4 (0xFF) is a placeholder for target channel number
            switch_message = [0x10, slot_id, 0x15, 0x1b, 0xFF, 0x00, 0x00]
            max_channels = 2
        elif dev_type.lower() == 'MX Master 3'.lower():
            # byte 6 is placeholder for key value
            # To be verified, probably wrong
            switch_detect_message = [0x11, slot_id, 0x08, 0x20, 0x00, 0xFF, 0x01]
            # To be verified, probably wrong
            easy_switch_keys = [0xD1, 0xD2, 0xD3]
            # To be verified
            # byte 2 is feature number for "CHANGE HOST". It can be checked in solaar by listing devices (solaar show)
            # byte 3 is magic number, for now, you simply have to check values until you hit correct one
            # byte 4 (0xFF) is a placeholder for target channel number
            switch_message = [0x10, slot_id, 0x0A, 0x11, 0xFF, 0x00, 0x00]
            max_channels = 3
        elif dev_type.lower() == 'MX Vertical'.lower():
            switch_message = [0x10, slot_id, 0x0C, 0x1C, 0xFF, 0x00, 0x00]
            easy_switch_keys = []
            switch_detect_message = []
            max_channels = 3
        elif dev_type.lower() == 'Ergo K860'.lower():
            switch_message = [0x10, slot_id, 0x09, 0x1C, 0xFF, 0x00, 0x00]
            # byte 6 is placeholder for key value
            # To be verified, probably wrong
            switch_detect_message = [0x11, slot_id, 0x08, 0x20, 0x00, 0xFF, 0x01]
            # To be verified, probably wrong
            easy_switch_keys = [0xD1, 0xD2, 0xD3]
            max_channels = 3
        elif dev_type.lower() == 'MK850'.lower():
            switch_message = [0x10, slot_id, 0x08, 0x11, 0xFF, 0x00, 0x00]
            # byte 6 is placeholder for key value
            # To be verified, probably wrong
            switch_detect_message = [0x11, slot_id, 0x08, 0x20, 0x00, 0xFF, 0x01]
            # To be verified, probably wrong
            easy_switch_keys = [0xD1, 0xD2, 0xD3]
            max_channels = 3
        elif dev_type.lower() == 'M720'.lower():
            switch_message = [0x10, slot_id, 0x09, 0x11, 0xFF, 0x00, 0x00]
            easy_switch_keys = []
            switch_detect_message = []
            max_channels = 3
        else:
            raise ValueError("Invalid Unifying device type")
        return UnifyingDevice(dev_type, slot_id, switch_detect_message, easy_switch_keys, switch_message, max_channels)

    def decode_target_channel_number(self, input_bytes):
        # Initial value indicating fault
        target_channel = -1
        if len(self.switch_detect_message) <= len(input_bytes):
            if (self.dev_type.lower() == 'MX Keys'.lower()) or (self.dev_type.lower() == 'MX Master 3'.lower()) or (self.dev_type.lower() == 'Ergo K860'.lower()) or (self.dev_type.lower() == 'MK850'.lower()):
                # Compare first 5 bytes (0 - 4) and byte 6. Byte 5 contains information about new channel.
                if (input_bytes[:4] == self.switch_detect_message[:4]) and \
                        (input_bytes[6] == self.switch_detect_message[6]):
                    for esk in self.easy_switch_keys:
                        if esk == input_bytes[5]:
                            target_channel = self.easy_switch_keys.index(esk)
            elif self.dev_type.lower() == 'MX Ergo'.lower() or len(self.switch_detect_message) == 0:
                # Device doesn't send information about switch button event
                target_channel = -2
            else:
                raise NotImplementedError
        return target_channel

    def switch_channel(self, channel_number):
        if channel_number < self.max_channels:
            self.switch_message[4] = channel_number
            logging.debug("Switching Unifying device \'%s\' at slot %d to channel %d" % (self.dev_type, self.slot_id,
                                                                                         channel_number))
            unifying_write(self.switch_message)
        else:
            logging.warning("Device \'%s\' does not support more than %d channels" % (self.dev_type, self.max_channels))

    def encode(self):
        return self.__dict__


class Monitor:
    def __init__(self, channel_to_monitor_id, channel_to_input_dict, vcp_message_number=60):
        self.channel_to_monitor_id = channel_to_monitor_id
        self.vcp_message_number = vcp_message_number
        self.channel_to_input_dict = channel_to_input_dict

    def switch_input(self, channel_number, self_channel):
        monitor = monitorcontrol.get_monitors()[self.channel_to_monitor_id[str(self_channel)]]
        with monitor:
            if channel_number < len(self.channel_to_input_dict) and channel_number < len(self.channel_to_monitor_id):
                old_input = "InputSource." + str(monitor.get_input_source())
                if old_input != self.channel_to_input_dict[str(channel_number)]:
                    logging.debug("Switch monitor %d from input %s to input %s" % (self.channel_to_monitor_id[str(self_channel)],
                                                                                   old_input,
                                                                                   self.channel_to_input_dict[str(channel_number)]))
                    monitor.set_input_source(self.channel_to_input_dict[str(channel_number)])
                else:
                    logging.debug("Monitor % is already set to input %d" % (self.channel_to_monitor_id[str(self_channel)],
                                                                            self.channel_to_input_dict[str(channel_number)]))
            else:
                logging.warning("Monitor is not configured for %d channels" % (channel_number + 1))

    def encode(self):
        return self.__dict__


class Config:
    def __init__(self, monitors, unifying_devices, self_channel):
        self.monitors = monitors
        self.unifying_devices = unifying_devices
        self.unifying_channel = self_channel

    def switch_channel(self, channel_number, detection_slot):
        for monitor in self.monitors:
            monitor.switch_input(channel_number, self.unifying_channel)

        for unifying_device in self.unifying_devices:
            if unifying_device.slot_id != detection_slot:
                unifying_device.switch_channel(channel_number)
            else:
                logging.debug("Skipping Unifying device \'%s\' at slot %d as the source of switch" %
                              (unifying_device.dev_type, unifying_device.slot_id))

    def encode(self):
        return self.__dict__


def populate_devices(self_channel):
    # Unifying receiver slots are numbered from 1
    mx_ergo = UnifyingDevice.get_from_type(dev_type='MX Ergo', slot_id=2)
    mx_keys = UnifyingDevice.get_from_type(dev_type='MX Keys', slot_id=1)
    # Main monitor
    main_monitor = Monitor(channel_to_monitor_id={'0': 2, '1': 1},
                           channel_to_input_dict={'0': 'HDMI2', '1': 'HDMI1'})

    config = Config(unifying_devices=[mx_keys, mx_ergo],
                    monitors=[main_monitor],
                    self_channel=self_channel)

    j = json.dumps(config, default=lambda o: o.encode(), indent=4)
    return config


def main_loop(self_channel, config_file):
    if config_file is not None:
        jconf = json.loads(config_file.read())
        unifying_devices = []
        monitors = []
        for jdev in jconf['unifying_devices']:
            if ('switch_detect_message' not in jdev) or \
                    ('easy_switch_keys' not in jdev) or \
                    ('switch_message' not in jdev) or \
                    ('max_channels' not in jdev):
                dev = UnifyingDevice.get_from_type(dev_type=jdev['dev_type'], slot_id=jdev['slot_id'])
            else:
                dev = UnifyingDevice(dev_type=jdev['dev_type'],
                                     slot_id=jdev['slot_id'],
                                     switch_detect_message=jdev['switch_detect_message'],
                                     easy_switch_keys=jdev['easy_switch_keys'],
                                     switch_message=jdev['switch_message'],
                                     max_channels=jdev['max_channels'])
            unifying_devices.append(dev)
        for jmonitor in jconf['monitors']:
            monitor = Monitor(channel_to_monitor_id=jmonitor['channel_to_monitor_id'],
                              channel_to_input_dict=jmonitor['channel_to_input_dict'],
                              vcp_message_number=jmonitor['vcp_message_number'])
            monitors.append(monitor)
        config = Config(monitors=monitors, unifying_devices=unifying_devices, self_channel=self_channel)



    else:
        config = populate_devices(self_channel)
    # Find the keyboard device (the one with easy_switch_keys or switch_detect_message)
    keyboard_slot = None
    for dev in config.unifying_devices:
        if len(dev.easy_switch_keys) > 0 or len(dev.switch_detect_message) > 0:
            keyboard_slot = dev.slot_id
            break

    PING_INTERVAL = 0.3  # seconds between keyboard pings
    PING_MISS_THRESHOLD = 2  # consecutive misses before declaring keyboard gone
    keyboard_present = True
    ping_miss_count = 0
    last_ping_time = time.time()

    logging.info("Using ping-based keyboard detection (slot %s, interval %.1fs, threshold %d)" %
                 (keyboard_slot, PING_INTERVAL, PING_MISS_THRESHOLD))

    while True:
        try:
            # Read any pending HID data (non-blocking with short timeout)
            read_bytes = unifying_listen(timeout_ms=200)
            if read_bytes:
                logging.debug("Raw HID data: %s" % str(read_bytes))

            # Detect keyboard connection via register 0x04 (if receiver sends it)
            if read_bytes and len(read_bytes) >= 4 and read_bytes[2] == 0x04 and keyboard_slot is not None and read_bytes[1] == keyboard_slot:
                is_connect = (read_bytes[3] & 0x40) == 0
                if is_connect:
                    if not keyboard_present:
                        keyboard_present = True
                        ping_miss_count = 0
                        logging.info("Keyboard connected (slot %d), switching mouse to self channel %d" % (keyboard_slot, self_channel))
                        for dev in config.unifying_devices:
                            if dev.slot_id != keyboard_slot:
                                dev.switch_channel(self_channel)
                continue

            # Original Easy-Switch key detection (fallback)
            if read_bytes:
                for unifying_device in config.unifying_devices:
                    channel_number = unifying_device.decode_target_channel_number(read_bytes)
                    if channel_number >= 0:
                        if channel_number != self_channel:
                            logging.info("Switch to channel " + str(channel_number) + " from device \'" +
                                         unifying_device.dev_type + "\' at slot " + str(unifying_device.slot_id))
                            config.switch_channel(channel_number, unifying_device.slot_id)

            # Periodic keyboard ping to detect departure
            now = time.time()
            if keyboard_slot is not None and (now - last_ping_time) >= PING_INTERVAL:
                last_ping_time = now
                is_alive = ping_device(keyboard_slot)
                if is_alive:
                    if not keyboard_present:
                        # Keyboard came back
                        keyboard_present = True
                        ping_miss_count = 0
                        logging.info("Keyboard returned (slot %d), switching mouse to self channel %d" % (keyboard_slot, self_channel))
                        for dev in config.unifying_devices:
                            if dev.slot_id != keyboard_slot:
                                dev.switch_channel(self_channel)
                    else:
                        ping_miss_count = 0
                else:
                    ping_miss_count += 1
                    logging.debug("Keyboard ping miss %d/%d" % (ping_miss_count, PING_MISS_THRESHOLD))
                    if ping_miss_count >= PING_MISS_THRESHOLD and keyboard_present:
                        # Keyboard is gone - switch mouse to other channel
                        keyboard_present = False
                        other_channel = 1 - self_channel
                        logging.info("Keyboard departed (slot %d), switching mouse to channel %d" % (keyboard_slot, other_channel))
                        for dev in config.unifying_devices:
                            if dev.slot_id != keyboard_slot:
                                dev.switch_channel(other_channel)
        except Exception as e:
            logging.error("Error in main loop: %s" % str(e))
            time.sleep(1)


if __name__ == '__main__':
    logFormatter = logging.Formatter("%(asctime)s,%(msecs)03d %(levelname)s %(message)s", datefmt='%Y-%m-%d %H:%M:%S')
    logFile = "log.log"
    fileHandler = RotatingFileHandler(logFile, mode='a', maxBytes=16 * 1024 * 1024, backupCount=2,
                                      encoding=None, delay=False)
    fileHandler.setFormatter(logFormatter)
    fileHandler.setLevel(logging.DEBUG)

    consoleHandler = logging.StreamHandler(sys.stdout)
    consoleHandler.setFormatter(logFormatter)

    app_log = logging.getLogger('root')
    app_log.setLevel(logging.DEBUG)

    app_log.addHandler(fileHandler)
    app_log.addHandler(consoleHandler)

    parser = argparse.ArgumentParser(description="Logitech unifying and HDMI monitor input switch")
    parser.add_argument('channel', metavar='CHANNEL', type=int,
                        help='Logitech Unifying self channel (this computer), numbered from  0')
    parser.add_argument('--config', '-c', type=argparse.FileType('r'), help="Configuration file")

    args = vars(parser.parse_args())

    logging.info("Discovering USB devices")
    usb_discover()

    logging.info("Opening persistent HID handles")
    open_hid_handles()

    # Enable wireless device connect/disconnect notifications on the receiver
    # HID++ 1.0: SET register 0x00, byte 0 bit 2 = wireless device status
    logging.info("Enabling wireless device notifications on receiver")
    unifying_write([0x10, 0xFF, 0x80, 0x00, 0x04, 0x00, 0x00])

    logging.info("Self unifying channel is " + str(args['channel']))
    main_loop(args['channel'], args['config'])
