import csv
import os
import sys
import time
import queue
import socket
import threading
import subprocess
from dataclasses import dataclass
import re

# Third-party
from dotenv import load_dotenv
import psutil
import pydirectinput
import pyperclip

# Optional Windows-only focus helpers
try:
    import win32gui
    import win32con
    import win32api
    import win32process
    import ctypes
except Exception:
    win32gui = None  # will fallback

# OBS websocket client (supports both legacy and new)
try:
    # Newer OBS (5.x protocol)
    from obswebsocket import obsws, requests as obsreq
except Exception:
    obsws = None
    obsreq = None


from obswebsocket import obsws, requests


# ------------- Config -------------
@dataclass
class Config:
    steam_exe: str
    dota_app_id: str
    dota_exe: str
    dota_log_file: str
    dota_window_title: str

    replays_csv: str
    replays_dir: str

    obs_host: str
    obs_port: int
    obs_password: str

    obs_scene_between: str
    obs_scene_live: str

    twitch_nick: str
    twitch_oauth: str
    twitch_channel: str
    twitch_cmd_prefix: str

    focus_retries: int
    focus_retry_seconds: float
    post_playdemo_wait: float
    console_key: str
    key_send_delay: float
    between_games_hold: float


# ------------- Utilities -------------


def getenv_default(key: str, default: str = "") -> str:
    val = os.getenv(key)
    return val if val is not None and val != "" else default


def load_config() -> Config:
    load_dotenv()
    return Config(
        steam_exe=getenv_default("STEAM_EXE"),
        dota_app_id=getenv_default("DOTA_APP_ID", "570"),
        dota_exe=getenv_default("DOTA_EXE"),
        dota_log_file=getenv_default("DOTA_LOG_FILE"),
        dota_window_title=getenv_default("DOTA_WINDOW_TITLE", "Dota 2"),
        replays_csv=getenv_default("REPLAYS_CSV", "./replays.csv"),
        replays_dir=getenv_default("REPLAYS_DIR", "/replays"),
        obs_host=getenv_default("OBS_HOST", "localhost"),
        obs_port=int(getenv_default("OBS_PORT", "4455")),
        obs_password=getenv_default("OBS_PASSWORD"),
        obs_scene_between=getenv_default("OBS_SCENE_BETWEEN", "Between Games"),
        obs_scene_live=getenv_default("OBS_SCENE_LIVE", ""),
        twitch_nick=getenv_default("TWITCH_NICK"),
        twitch_oauth=getenv_default("TWITCH_OAUTH"),
        twitch_channel=getenv_default("TWITCH_CHANNEL"),
        twitch_cmd_prefix=getenv_default("TWITCH_CMD_PREFIX", "!"),
        focus_retries=int(getenv_default("FOCUS_RETRIES", "15")),
        focus_retry_seconds=float(getenv_default("FOCUS_RETRY_SECONDS", "1.0")),
        post_playdemo_wait=float(getenv_default("POST_PLAYDEMO_WAIT", "6.0")),
        console_key=getenv_default("CONSOLE_KEY", "F11"),
        key_send_delay=float(getenv_default("KEY_SEND_DELAY", "0.08")),
        between_games_hold=float(getenv_default("BETWEEN_GAMES_HOLD", "4.0")),
    )


# ------------- OBS Client -------------
class OBSClient:
    def __init__(self, host: str, port: int, password: str):
        self.host, self.port, self.password = host, port, password
        self.ws = None

    def connect(self):
        if obsws is None:
            print("[OBS] obs-websocket-py not installed; OBS features disabled.")
            return
        try:
            self.ws = obsws(self.host, self.port, self.password)
            self.ws.connect()
            print("[OBS] Connected")
        except Exception as e:
            print(f"[OBS] Failed to connect: {e}")
            self.ws = None

    def safe_set_scene(self, scene: str):
        if not scene or self.ws is None:
            return
        try:
            # OBS WebSocket 5.x+ uses camelCase and SetCurrentProgramScene
            self.ws.call(obsreq.SetCurrentProgramScene(sceneName=scene))
            print(f"[OBS] Switched to scene: {scene}")
        except Exception as e:
            try:
                # Fallback for legacy versions
                self.ws.call(obsreq.SetCurrentScene({"scene-name": scene}))
                print(f"[OBS] Switched (legacy) to scene: {scene}")
            except Exception as e2:
                print(f"[OBS] Failed to switch scene '{scene}': {e2}")

    def disconnect(self):
        if self.ws is not None:
            try:
                self.ws.disconnect()
            except Exception:
                pass
            self.ws = None
            print("[OBS] Disconnected")


# ------------- Twitch IRC Listener -------------
class TwitchListener(threading.Thread):
    def __init__(
        self,
        nick: str,
        oauth: str,
        channel: str,
        cmd_prefix: str,
        command_queue: queue.Queue,
    ):
        super().__init__(daemon=True)
        self.nick = nick
        self.oauth = oauth
        self.channel = channel.lower()
        self.cmd_prefix = cmd_prefix
        self.command_queue = command_queue
        self.sock = None
        self.running = True

    def run(self):
        try:
            self.sock = socket.socket()
            self.sock.connect(("irc.chat.twitch.tv", 6667))
            self.sock.send(f"PASS {self.oauth}\r\n".encode("utf-8"))
            self.sock.send(f"NICK {self.nick}\r\n".encode("utf-8"))
            self.sock.send(f"JOIN #{self.channel}\r\n".encode("utf-8"))
            print(f"[Twitch] Connected as {self.nick} to #{self.channel}")
        except Exception as e:
            print(f"[Twitch] Connection failed: {e}")
            return

        buffer = ""
        while self.running:
            try:
                data = self.sock.recv(4096)
                if not data:
                    break
                buffer += data.decode("utf-8", errors="ignore")
                while "\r\n" in buffer:
                    line, buffer = buffer.split("\r\n", 1)
                    if line.startswith("PING"):
                        # Keep-alive
                        self.sock.send("PONG :tmi.twitch.tv\r\n".encode("utf-8"))
                        continue
                    # Parse messages: :user!user@user.tmi.twitch.tv PRIVMSG #channel :message
                    if "PRIVMSG" in line:
                        try:
                            prefix, msg = line.split(" PRIVMSG ", 1)
                            user = prefix.split("!")[0][1:]
                            text = msg.split(" :", 1)[1]
                            self._handle_message(user, text)
                        except Exception:
                            pass
            except Exception:
                time.sleep(1)

    def _handle_message(self, user: str, text: str):
        text = text.strip()
        if not text.startswith(self.cmd_prefix):
            return
        cmd = text[len(self.cmd_prefix) :].lower()
        if cmd in {"p1", "p2", "p3", "p4", "p5", "p6", "p7", "p8", "p9", "p10"}:
            self.command_queue.put(cmd)
            print(f"[Twitch] Command from {user}: {cmd}")

    def stop(self):
        self.running = False
        try:
            if self.sock:
                self.sock.close()
        except Exception:
            pass


# ------------- Dota Process & Window -------------


def is_process_running(name_contains: str) -> bool:
    name_contains_lower = name_contains.lower()
    for p in psutil.process_iter(["name"]):
        try:
            if p.info["name"] and name_contains_lower in p.info["name"].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False


def launch_dota(cfg: Config):
    dota_args = ["-console", "-novid", "-condebug", "-gamestateintegration"]

    if cfg.dota_exe:
        print("[Dota] Launching via DOTA_EXE…")
        try:
            subprocess.Popen([cfg.dota_exe] + dota_args)
        except Exception as e:
            print(f"[Dota] Failed to launch dota_exe: {e}")
    elif cfg.steam_exe:
        print("[Dota] Launching via Steam…")
        try:
            subprocess.Popen([cfg.steam_exe, "-applaunch", cfg.dota_app_id] + dota_args)
        except Exception as e:
            print(f"[Dota] Failed to launch Steam: {e}")
    else:
        print("[Dota] No launch path configured (set STEAM_EXE or DOTA_EXE).")


def _set_foreground(hwnd):
    if hwnd and win32gui:
        try:
            # Force foreground even if another process has focus
            ctypes.windll.user32.ShowWindow(hwnd, 5)  # SW_SHOW
            ctypes.windll.user32.SetForegroundWindow(hwnd)
        except Exception:
            pass


def focus_dota_window(cfg: Config) -> bool:
    if not win32gui:
        print("[Focus] win32gui not available; cannot reliably focus Dota.")
        return False
    hwnd = win32gui.FindWindow(None, cfg.dota_window_title)
    if hwnd == 0:
        # try partial title match
        def _enum_handler(h, result):
            title = win32gui.GetWindowText(h)
            if cfg.dota_window_title.lower() in title.lower():
                result.append(h)

        result = []
        win32gui.EnumWindows(_enum_handler, result)
        hwnd = result[0] if result else 0

    if hwnd:
        _set_foreground(hwnd)
        return True
    return False


# ------------- Keystroke Helpers -------------


def press_key(key: str, delay: float = 0.08):
    pydirectinput.press(key)
    time.sleep(delay)


def send_console_commands(cfg, commands: list[str]):
    # Open console
    press_key(cfg.console_key, cfg.key_send_delay)

    for cmd in commands:
        # Copy command to clipboard
        pyperclip.copy(cmd)

        # Press Ctrl+V manually
        pydirectinput.keyDown("ctrl")
        pydirectinput.press("v")
        pydirectinput.keyUp("ctrl")

        # Press Enter to execute
        press_key("enter", cfg.key_send_delay)
        time.sleep(0.05)

    # Close console
    press_key(cfg.console_key, cfg.key_send_delay)


# ------------- Log Tailer -------------
class LogTailer(threading.Thread):
    def __init__(self, logfile: str, end_event: threading.Event):
        super().__init__(daemon=True)
        self.logfile = logfile
        self.end_event = end_event
        self.running = True

    def run(self):
        if not self.logfile:
            print("[Log] No DOTA_LOG_FILE set; replay end detection disabled.")
            return
        try:
            with open(self.logfile, "r", encoding="utf-8", errors="ignore") as f:
                # seek to end
                f.seek(0, os.SEEK_END)
                while self.running:
                    line = f.readline()
                    if not line:
                        time.sleep(0.2)
                        continue
                    if re.search(r"GameEnd\s*$", line):
                        print("[Log] Detected GameEnd → replay end")
                        self.end_event.set()
        except FileNotFoundError:
            print(f"[Log] File not found: {self.logfile}")
        except Exception as e:
            print(f"[Log] Tailer error: {e}")

    def stop(self):
        self.running = False


# ------------- Replay Runner -------------
class ReplayRunner:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.obs = OBSClient(cfg.obs_host, cfg.obs_port, cfg.obs_password)
        self.twitch_queue = queue.Queue()
        self.twitch = TwitchListener(
            cfg.twitch_nick,
            cfg.twitch_oauth,
            cfg.twitch_channel,
            cfg.twitch_cmd_prefix,
            self.twitch_queue,
        )
        self.replay_end = threading.Event()
        self.log_tailer = LogTailer(cfg.dota_log_file, self.replay_end)
        self.last_command_time = 0  # Track last execution time
        self.command_cooldown = 60  # Cooldown in seconds

    def load_replay_ids(self) -> list[str]:
        ids = []

        # Read existing IDs from the CSV
        with open(self.cfg.replays_csv, newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            for row in reader:
                if not row:
                    continue
                for cell in row:
                    cell = cell.strip()
                    if not cell:
                        continue
                    # keep only digits
                    mid = "".join(ch for ch in cell if ch.isdigit())
                    if mid:
                        ids.append(mid)

        print(f"[Replay] Loaded {len(ids)} ids")

        if not ids:
            return []

        # Let's say you want to process and REMOVE the first replay
        replay_to_remove = ids[0]

        # Filter out the replay we just processed
        remaining_ids = [rid for rid in ids if rid != replay_to_remove]

        # Write the remaining IDs back to the CSV
        with open(self.cfg.replays_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for rid in remaining_ids:
                writer.writerow([rid])

        print(f"[Replay] Removed replay ID: {replay_to_remove}")
        return ids

    def ensure_dota_ready(self):
        if not is_process_running("dota"):
            launch_dota(self.cfg)
            # wait a bit for process
            for _ in range(60):
                if is_process_running("dota"):
                    break
                time.sleep(1)
            time.sleep(10)  # extra wait for Dota to initialize so console works
        # focus window
        for i in range(self.cfg.focus_retries):
            if focus_dota_window(self.cfg):
                print("[Focus] Dota window focused")
                return True
            time.sleep(self.cfg.focus_retry_seconds)
        print("[Focus] Failed to focus Dota window")
        return False

    def handle_twitch_commands(self):
        # Drain queue, execute last relevant player switch
        key_to_press = None

        current_time = time.time()
        if current_time - self.last_command_time < self.command_cooldown:
            # Too soon to execute another command
            return

        while True:
            try:
                cmd = self.twitch_queue.get_nowait()
            except queue.Empty:
                break

            if cmd.startswith("p") and cmd[1:].isdigit():
                n = int(cmd[1:])
                # In Dota spectator: 1..9 and 0 for 10
                key_to_press = "0" if n == 10 else str(n)

        if key_to_press:
            # Ensure focus before key
            focus_dota_window(self.cfg)
            press_key(key_to_press, self.cfg.key_send_delay)
            print(f"[Twitch] Switched player via key '{key_to_press}'")

            # Update last execution time
            self.last_command_time = current_time

    def show_between_games(self):
        self.obs.safe_set_scene(self.cfg.obs_scene_between)
        time.sleep(self.cfg.between_games_hold)

    def show_live(self):
        if self.cfg.obs_scene_live:
            self.obs.safe_set_scene(self.cfg.obs_scene_live)

    def play_replay(self, match_id: str):
        # Ensure focus
        self.ensure_dota_ready()
        # Commands sequence
        replay_file_name = f"{match_id}.dem"
        commands = [
            "disconnect",
            f"playdemo {self.cfg.replays_dir}/{replay_file_name}",
            "dota_spectator_mode 3",
        ]
        # send commands
        send_console_commands(self.cfg, commands)
        time.sleep(self.cfg.post_playdemo_wait)
        # Set spectator cam mode 3
        self.show_live()

        # wait until replay end signal
        self.replay_end.clear()
        while not self.replay_end.is_set():
            self.handle_twitch_commands()
            time.sleep(0.2)
        print(f"[Replay] Finished match {match_id}")

    def run(self):
        # Connect services
        self.obs.connect()
        self.twitch.start()
        self.log_tailer.start()

        try:
            while True:
                replay_ids = (
                    self.load_replay_ids()
                )  # reads current CSV (which may have changed)
                if not replay_ids:
                    print("[Replay] No IDs found — exiting.")
                    break

                # Always take the next available ID from the freshly loaded list
                mid = replay_ids[0]

                self.show_between_games()
                self.play_replay(mid)
        finally:
            # Cleanup
            self.twitch.stop()
            self.log_tailer.stop()
            self.obs.disconnect()
            print("[Runner] Done.")


def main():
    if not sys.platform.startswith("win"):
        print(
            "[WARN] This script is tailored for Windows. Key injection and window focusing may differ on other OSes."
        )
    cfg = load_config()
    runner = ReplayRunner(cfg)
    runner.run()


if __name__ == "__main__":
    main()
