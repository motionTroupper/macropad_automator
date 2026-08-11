import usb_cdc
import storage
import os

usb_cdc.enable(console=True, data=True)    # Enable console and data

try:
    os.stat('usb_exposed')
    # Flag exists: expose USB drive for editing
except:
    # Default: hide USB drive (security)
    storage.disable_usb_drive()