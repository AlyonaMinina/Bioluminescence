# 🤖 Imaging System Controller

This folder contains the Raspberry Pi–based imaging controller and locally hosted web GUI for the imaging system.

The system is built around:
- Nikon D800 camera
- Raspberry Pi 4
- Flask-based local web application

---

## ✨ Features

- Live preview for camera focusing
- Acquisition of:
  - day images
  - bioluminescence images
- File preview and download through the web GUI
- Automatic timestamping of images to prevent overwriting
- Currently supports acquisition of:
  - 1 day image
  - 1 bioluminescence image
- Timelapse acquisition is not yet implemented

---

## 🫖 Running the GUI as a Background Service

To allow the GUI to start automatically and run continuously on the Raspberry Pi:

### 1. Copy the `.service` file to the Raspberry Pi

Example using `scp`:

```bash
scp controller.service pi@<RPI_IP>:~
```

---

### 2. Install and enable the service on the Raspberry Pi

Run the following commands on the Raspberry Pi:

```bash
sudo cp ~/controller.service /etc/systemd/system/

sudo systemctl daemon-reload

sudo systemctl enable controller

sudo systemctl start controller
```

---

## 🌐 Accessing the Web GUI

1. Connect a laptop to the Raspberry Pi hotspot/network.
2. Determine the Raspberry Pi IP address:

```bash
hostname -I
```

3. Open a web browser and navigate to:

```text
http://<RPI_IP>:5000/
```

Example:

```text
http://10.42.0.1:5000/
```

---

## 🦾 Restarting the Controller During Development

After updating the GUI/controller code, restart the service:

```bash
sudo systemctl stop controller

sleep 2

pkill -f gphoto2

sleep 1

# Copy/update the new controller files here

sudo systemctl start controller
```
