import datetime
import json
import os
import time
import traceback
import requests
import logging
from datetime import date
from PIL import Image, ImageDraw, ImageFont
from obswebsocket import obsws, requests as obsrequests
import pyautogui
import socket
import re
from collections import Counter
import random
import pydirectinput
from dotenv import load_dotenv
import csv
import threading
import pyperclip

today = date.today()
formatted_date = today.strftime("%Y-%m-%d")
logging.basicConfig(
    filename=f"{formatted_date}.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

load_dotenv()

DOTA2_CLIENT_PATH = os.getenv("DOTA2_CLIENT_PATH")
REPLAY_PATH = os.getenv("REPLAY_PATH")
REPLAY_CSV = os.getenv("REPLAY_CSV")
ISDEBUG = os.getenv("ISDEBUG")
OBS_WEBSOCKET_PASSWORD = os.getenv("OBS_WEBSOCKET_PASSWORD")
TWITCHTOKEN = os.getenv("TWITCHTOKEN")
TWITCHCHANNELNAME = os.getenv("TWITCHCHANNELNAME")
DOTALOGFILE = os.getenv("DOTALOGFILE")

# OBS SCENES
DOTA2CLIENTSCENE = "Dota2Client"
BETWEENGAMESSCENE = "BetweenGames"

# OBS INPUT
LOADINGSCREENIMAGEINPUT = "LoadingScreenImage"

# BACKGROUND IMAGE
BACKGROUNDIMAGE = "background_white.jpg"

if (
    not DOTA2_CLIENT_PATH
    or not REPLAY_PATH
    or not REPLAY_CSV
    or not ISDEBUG
    or not OBS_WEBSOCKET_PASSWORD
    or not TWITCHTOKEN
    or not TWITCHCHANNELNAME
):
    raise Exception("You need to set local all local variables. See .env.dist")
_next_match = True
_lock = threading.Lock()


def main():
    try:
        matches = []

        with open("heroData.json", "r") as file:
            heroes = json.load(file)

        log_thread = threading.Thread(target=read_logs, daemon=True)
        log_thread.start()

        while True:
            matches = get_all_match_details(matches)
            if get_next_match():
                set_next_match(False)
            else:
                time.sleep(1)
                continue

            image_path = generate_loadscreen_image(matches, heroes)
            votedMatch = take_votes(image_path) - 1
            match = matches[votedMatch]

            match_id = match["match_id"]
            replay_file_name = f"{match_id}.dem"

            if not replay_file_name:
                continue

            enter_console_command(
                [
                    "disconnect",
                    f"playdemo /replays/{replay_file_name}",
                    "dota_spectator_mode 3",
                ]
            )
            ws = get_obs_websocket()
            set_current_preview_scene(ws, DOTA2CLIENTSCENE)
            trigger_studio_mode_transition(ws)

            ws.disconnect()
            remove_entry_from_csv(votedMatch)

    except Exception as ex:
        logger.error(f"An error occurred: {ex}\n{traceback.format_exc()}")


def get_next_match():
    with _lock:
        return _next_match


def set_next_match(value):
    global _next_match
    with _lock:
        _next_match = value


def enter_console_command(consoleCommands):
    pydirectinput.press("f11")
    if isinstance(consoleCommands, list):
        for consoleCommand in consoleCommands:
            pyautogui.write(consoleCommand)
            pydirectinput.press("enter")
            time.sleep(1)
    else:
        pyautogui.write(consoleCommands)
        pydirectinput.press("enter")

    pydirectinput.press("f11")


def read_logs():
    with open(DOTALOGFILE, "r") as file:
        file.seek(0, 2)
        while True:
            line = file.readline()
            if not line:
                time.sleep(0.1)
                continue
            if line.startswith("COMBAT SUMMARY"):
                set_next_match(True)


def remove_entry_from_csv(entry_to_remove):
    # Read the CSV file and store its contents in memory
    with open(REPLAY_CSV, "r", newline="") as file:
        reader = csv.reader(file)
        rows = list(reader)

    if ISDEBUG:
        return

    try:
        index_to_remove = int(entry_to_remove)
        if index_to_remove < 0 or index_to_remove >= len(rows):
            logger.error("Invalid row index to remove.")
            return
        rows.pop(index_to_remove)
    except ValueError:
        logger.error("Invalid row index to remove. Please provide a valid integer.")
        return

    # Write the modified data back to the CSV file
    with open(REPLAY_CSV, "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerows(rows)

    logger.debug("Entry removed successfully.")


def get_all_match_details(matches):
    match_ids_from_array = [match["match_id"] for match in matches]

    with open(REPLAY_CSV, mode="r") as file:
        csv_reader = csv.reader(file)
        my_list = [row[0] for row in csv_reader if row]

    missing_match_ids = set(my_list) - set(match_ids_from_array)

    for matchId in missing_match_ids:
        matchUrl = f"https://api.opendota.com/api/matches/{matchId}"
        matchResponse = requests.get(matchUrl)

        if matchResponse.status_code != 200:
            logger.error(
                f"Error making get match details call {matchId}. Status code: {matchResponse.status_code}"
            )
            continue

        matches.append(matchResponse.json())

    return matches


def get_match_details():
    matchDetails = []
    with open(REPLAY_CSV, mode="r") as file:
        csv_reader = csv.reader(file)
        my_list = [row[0] for row in csv_reader if row]

    for matchId in my_list:
        matchUrl = f"https://api.opendota.com/api/matches/{matchId}"
        matchResponse = requests.get(matchUrl)

        if matchResponse.status_code != 200:
            logger.error(
                f"Error making get match details call {matchId}. Status code: {matchResponse.status_code}"
            )
            continue

        matchDetails.append(matchResponse.json())

    return matchDetails


def generate_loadscreen_image(matches, heroes):
    image = Image.open(BACKGROUNDIMAGE)
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype("arial.ttf", 16)
    matchesDrawn = 5
    if len(matches) < 5:
        matchesDrawn = len(matches)
    for x1 in range(matchesDrawn):
        players = matches[x1].get("players")
        player_hero_ids = [player["hero_id"] for player in players]
        filtered_heroes = [
            next(hero for hero in heroes if hero["id"] == hero_id)
            for hero_id in player_hero_ids
        ]

        top = 100
        left = 50 + (360 * x1)
        # Hero names
        for x2 in range(10):
            position = (left, top)

            draw.text(
                position, filtered_heroes[x2]["localized_name"], (0, 0, 0), font=font
            )

            if x2 == 4:
                top += 150
            else:
                top += 50

    top = 700
    for x3 in range(5):
        left = 50 + (360 * x3)
        position = (left, top)
        draw.text(position, f"!vote {x3 + 1}", (0, 0, 0), font=font)

    current_dir = os.getcwd()
    imageName = f"tmpImg_{round(time.time())}.png"
    image_path = os.path.join(current_dir, imageName)

    image.save(image_path)

    return image_path


def get_obs_websocket():
    ws = obsws("localhost", 4455, OBS_WEBSOCKET_PASSWORD)
    ws.connect()
    return ws


def set_current_preview_scene(ws, sceneName):
    ws.call(obsrequests.SetCurrentPreviewScene(sceneName=sceneName))


def trigger_studio_mode_transition(ws):
    ws.call(obsrequests.TriggerStudioModeTransition())


def set_input_settings(ws, inputName, inputSettings):
    ws.call(
        obsrequests.SetInputSettings(
            inputName=inputName, inputSettings=inputSettings, overlay=True
        )
    )


def get_irc_socket_object():
    server = "irc.chat.twitch.tv"
    port = 6667
    irc = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    irc.connect((server, port))
    irc.send(f"PASS {TWITCHTOKEN}\n".encode("utf-8"))
    irc.send(f"NICK {TWITCHCHANNELNAME}\n".encode("utf-8"))
    irc.send(f"JOIN #{TWITCHCHANNELNAME}\n".encode("utf-8"))
    irc.settimeout(1)
    return irc


def read_chat(process):
    irc = get_irc_socket_object()
    latestFollowingCommandTime = time.time() - 30

    while process.is_running() is True:
        try:
            # Get the chat line
            response = irc.recv(2048).decode("utf-8")

            # Respond to pings from the server, required to keep the connection alive
            if response.startswith("PING"):
                irc.send("PONG\n".encode("utf-8"))
                continue

            responsePattern = re.compile(r":(.*?)!.*?:([^:]+)")

            match = responsePattern.search(response)

            if match:
                logger.debug(f"Recieves chat message: {response}")
                name = match.group(1)
                message = match.group(2).strip()

                followCommand = re.search(r"(!f\s*|f\s+)(10|[1-9])", message)
                if followCommand and latestFollowingCommandTime < time.time() - 30:
                    number = f"{followCommand.group(2)}"
                    pydirectinput.press(number)
                    latestFollowingCommandTime = time.time()
            else:
                logger.debug(f"Could not read chat message: {response}")
        except socket.timeout:
            continue

    irc.close()


def take_votes(image_path):
    ws = get_obs_websocket()

    set_input_settings(ws, LOADINGSCREENIMAGEINPUT, {"file": image_path})
    set_current_preview_scene(ws, BETWEENGAMESSCENE)
    trigger_studio_mode_transition(ws)

    ws.disconnect()
    irc = get_irc_socket_object()
    irc.settimeout(1)
    sleepTime = 5 if ISDEBUG else 45

    start_time = time.time()
    end_time = start_time + sleepTime

    voteMessages = []

    while time.time() < end_time:
        try:
            # Get the chat line
            response = irc.recv(2048).decode("utf-8")

            # Respond to pings from the server, required to keep the connection alive
            if response.startswith("PING"):
                irc.send("PONG\n".encode("utf-8"))
                continue

            responsePattern = re.compile(r":(.*?)!.*?:([^:]+)")

            match = responsePattern.search(response)
            if match:
                logger.debug(f"Recieves chat message: {response}")
                name = match.group(1)
                message = match.group(2).strip()

                legitVote = re.search(r"!v([1-9])|!vote\s?([1-5])", message)
                if legitVote:
                    add_vote(voteMessages, name, message)
            else:
                logger.debug(f"Could not read chat message: {response}")
        except socket.timeout:
            continue

    irc.close()

    if not voteMessages:
        return random.randint(1, 5)

    votes = [item[1] for item in voteMessages]
    count = Counter(votes)
    return count.most_common(1)[0][0]


def add_vote(votes, name, vote):
    voteChanged = False
    for arr in votes:
        if arr[0] == name:
            arr[1] = vote
            voteChanged = True
            break
    if voteChanged == False:
        votes.append((name, extract_number(vote)))
    return votes


def extract_number(s):
    # Search for a number between 1 and 10
    match = re.search(r"\b(10|[1-9])\b", s)
    if match:
        return int(match.group(0))
    else:
        return 0


if __name__ == "__main__":
    main()

# pyinstaller --add-data "heroData.json;." --add-data "local_settings.py;." --add-data "background_white.jpg;."  main.py


# Available Input Kinds:
# image_source
# color_source
# slideshow
# browser_source
# ffmpeg_source
# text_gdiplus
# text_ft2_source
# vlc_source
# monitor_capture
# window_capture
# game_capture
# dshow_input
# wasapi_input_capture
# wasapi_output_capture
# wasapi_process_output_capture
