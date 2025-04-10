# Mullvad Server Tester

A graphical user interface (GUI) application for automatically testing the performance (ping and speed) of Mullvad VPN servers.

## Features

*   **Server Loading:** Loads Mullvad server data (requires `relays.json` cache file).
*   **Server Selection:**
    *   Displays servers in a sortable list (by hostname, city, country, status, ping, download/upload speeds).
    *   Filter servers by country.
    *   Select individual servers or test all visible servers.
*   **Automated Testing:**
    *   Connects sequentially to selected servers using the Mullvad CLI.
    *   Verifies the connection status.
    *   Performs a ping test (to 1.1.1.1) through the connected server.
    *   Runs a speed test using the Ookla Speedtest CLI through the connected server.
    *   Disconnects from the server.
*   **Results Display:** Shows status, ping (ms), download (Mbps), and upload (Mbps) for each tested server in the list.
*   **Test Control:** Start, stop, and pause the testing process.
*   **Export:** Export the test results to a CSV file.
*   **Configuration:** Settings menu to configure timeouts, ping count, delays, and theme.
*   **Theming:** Supports light, dark, and system themes (using `sv-ttk` if installed).
*   **Logging:** Logs detailed information to `~/mullvad_tester.log`.

## Dependencies

This application requires the following command-line tools to be installed and accessible in your system's PATH:

1.  **Mullvad VPN Client:** The official Mullvad client provides the `mullvad` CLI tool needed for connecting, disconnecting, and managing server locations. Download from [mullvad.net](https://mullvad.net/).
2.  **Ookla Speedtest CLI:** Used for performing the speed tests.
    *   **macOS (using Homebrew):**
        ```bash
        brew tap teamookla/speedtest
        brew update
        brew install speedtest --force
        ```
    *   **Other Systems:** Download from [speedtest.net/apps/cli](https://speedtest.net/apps/cli).
    *   **Important:** You might need to run `speedtest` once manually in your terminal after installation to accept the license agreement.

## Installation & Setup

1.  **Clone or Download:** Get the application code.
2.  **Install Python Dependencies:** (Assuming standard libraries and Tkinter are present with Python)
    *   Optional (for better theming): `pip install sv-ttk`
3.  **Mullvad Cache:** Ensure the Mullvad client has generated its server cache file (`relays.json`). The application will try to find it automatically, but you can specify a custom path in the Settings menu if needed. Common locations:
    *   macOS: `~/Library/Caches/net.mullvad.mullvad/relays.json`
    *   Linux: `~/.cache/Mullvad VPN/relays.json`
    *   Windows: `%LOCALAPPDATA%\Mullvad VPN\cache\relays.json`
4.  **Run the Application:**
    ```bash
    python automated_tester.py
    ```

## Usage

1.  Launch the application using `python automated_tester.py`.
2.  Wait for the server list to load.
3.  Optionally, filter the servers by selecting a country from the dropdown menu.
4.  Select servers to test by clicking the checkbox next to their hostname. Click the header checkbox to select/deselect all visible servers.
5.  Click "Start Test" (or "Start Test (X Selected)").
6.  The application will connect to each selected server, run tests, and display results.
7.  Use "Pause"/"Resume" or "Stop Test" to control the process.
8.  Double-click a server row to attempt a direct connection to that server (outside the testing sequence).
9.  Go to `File -> Export Results to CSV...` to save the current results.
10. Access `File -> Settings` to adjust testing parameters.

## Screenshot

![Application Screenshot](Screenshot.png)
