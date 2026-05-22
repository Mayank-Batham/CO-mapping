# CIA Marks Mapper - Android App (Zero Configuration Setup)

This directory contains the **Android Studio Project** for the **CIA Marks & CO Mapper**. The app uses native Android technologies (Kotlin and Coroutines) to wrap your Streamlit application inside a secure, high-performance WebView and features a **Zero-Configuration Automatic Subnet Discovery** system.

---

## 🚀 How It Works (Zero-Configuration)

Instead of forcing you or other users to figure out and type your laptop's local IP address (e.g., `192.168.1.45`), the app does it automatically:
1. **Detects Net Interface**: At startup, it reads the active Wi-Fi connection's IP address and determines the gateway subnet base (e.g., `192.168.1.`).
2. **Asynchronous Parallel Port Probe**: Spawns **254 lightweight Kotlin Coroutines** to concurrently ping port `8501` across all IP addresses on the network subnet.
3. **Immediate Connection**: 
   - The entire sweep completes in **under 1 second**.
   - As soon as the active laptop server is found, the app automatically transitions from the splash screen to the Streamlit UI.
4. **Fallback Options**:
   - If no network is present or the server is not on the Wi-Fi, it checks the standard Android Emulator gateway (`10.0.2.2:8501`).
   - If that also fails, it displays an elegant **Retry** layout with instructions to verify network connectivity.

---

## 🛠️ Step-by-Step Deployment Guide

Since only **one laptop** is running the server, follow these simple steps to deploy and run:

### Step 1: Run the Server on Your Laptop
1. Open PowerShell or Terminal in the project root directory (`c:\Users\mayan\Desktop\gotta`).
2. Activate your Python virtual environment and start the Streamlit server:
   ```powershell
   .\venv\Scripts\python -m streamlit run app.py
   ```
3. Keep the server running. Ensure that your laptop remains connected to your local Wi-Fi router.

### Step 2: Open and Compile the Android App
1. Launch **Android Studio**.
2. Click **Open** (or **File > Open**) and select the `android_mapper` folder inside this project directory.
3. Android Studio will automatically recognize the project structure, download the required Gradle build tools, and sync the dependencies.

### Step 3: Run the App on Your Device / Emulator
* **To run on a Physical Android Phone**:
  1. Enable **Developer Options** and **USB Debugging** on your phone.
  2. Connect your phone to your laptop using a USB cable. Make sure **both the laptop and the phone are connected to the exact same Wi-Fi network**.
  3. In Android Studio, select your phone from the device dropdown at the top and click the green **Run (Play)** button.
  4. The app will install and open. It will show a sleek splash screen saying `"Searching for laptop server..."` and then automatically load the Streamlit dashboard!

* **To run on the Android Studio Emulator**:
  1. Click the green **Run (Play)** button in Android Studio to launch the built-in emulator.
  2. The app will automatically connect to the laptop using the developer loopback (`10.0.2.2:8501`).

---

## 📂 Project Structure & Features Included

- **File Upload Support**: Customized Android `WebChromeClient` intercepts Streamlit file requests and invokes the native Android file picker (`ACTION_GET_CONTENT`), filtered for Excel files (`.xlsx`, `.xls`) and Word documents (`.docx`).
- **File Download Integration**: Built-in Android `DownloadManager` hooks into the WebView's download stream to download generated, consolidated Excel files directly into the Android device's **Downloads** folder, showing native system download notifications.
- **Orientation & Layout Stability**: Fully configured in `AndroidManifest.xml` (`orientation|screenSize` config changes) to prevent screen-rotation reloads and preserve user data.
- **Premium Styling & UI**: Standard Slate-Gray theme variables defined in `colors.xml` and beautiful progress/splash layout implemented in `activity_main.xml`.
