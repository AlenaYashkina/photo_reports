import os
import json
import time
import logging
from telethon.sync import TelegramClient
from telethon.tl.types import PeerChannel, InputMessagesFilterPhotos
from tqdm import tqdm

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

CONFIG_FILE = 'config.json'
DOWNLOADED_PHOTOS_FILE = 'downloaded_photos.json'


def load_config(file_path):
    """
    Load configuration from a JSON file.

    :param file_path: Path to the JSON configuration file.
    :return: Parsed configuration as a dictionary.
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        logger.error(f"Configuration file '{file_path}' not found.")
        exit(1)
    except json.JSONDecodeError as e:
        logger.error(f"Error parsing JSON in '{file_path}': {e}")
        exit(1)


def load_downloaded_photos(file_path):
    """
    Load already downloaded photos info from a JSON file.

    :param file_path: Path to the JSON file with downloaded photos data.
    :return: Dictionary with downloaded photo paths.
    """
    if os.path.exists(file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            logger.warning(f"Could not parse '{file_path}'. Starting with an empty dictionary.")
    return {}


def save_downloaded_photos(file_path, data):
    """
    Save the downloaded photos info to a JSON file.

    :param file_path: Path to the JSON file.
    :param data: Dictionary containing downloaded photo paths.
    """
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        logger.error(f"Error saving downloaded photos info: {e}")


def get_chat_and_topic_titles(chat_id, topic_id, client):
    """
    Retrieve the chat title and a formatted topic title.

    :param chat_id: Telegram chat/channel ID.
    :param topic_id: Topic identifier (reply-to message ID).
    :param client: Initialized TelegramClient instance.
    :return: Tuple (chat_title, topic_title)
    """
    logger.info(f"Fetching titles for chat ID {chat_id} and topic ID {topic_id}...")
    try:
        chat = client.get_entity(PeerChannel(chat_id))
        # Replace spaces with underscores for folder naming
        chat_title = chat.title.replace(" ", "_")
        topic_title = f"topic_{topic_id}"
        return chat_title, topic_title
    except Exception as e:
        logger.error(f"Error retrieving titles: {e}")
        return "unknown_chat", f"topic_{topic_id}"


def download_photos(client, phone_number, chat_id, topics, downloaded_photos):
    """
    Download photos from specified Telegram topics.

    :param client: TelegramClient instance.
    :param phone_number: Phone number for authorization.
    :param chat_id: Chat/channel ID.
    :param topics: List of topic IDs.
    :param downloaded_photos: Dictionary with already downloaded photos.
    """
    try:
        client.connect()

        # Authorize user if not already authorized
        if not client.is_user_authorized():
            logger.info("Authorizing user...")
            client.send_code_request(phone_number)
            code = input("Enter the code: ")
            client.sign_in(phone_number, code)

        # Get chat entity and format chat title
        chat = client.get_entity(PeerChannel(chat_id))
        chat_title = chat.title.replace(" ", "_")

        for topic_id in topics:
            # Get titles for chat and topic
            _, topic_title = get_chat_and_topic_titles(chat_id, topic_id, client)
            logger.info(f"Chat title: {chat_title}, Topic title: {topic_title}")

            # Prepare the directory to save photos
            save_path = os.path.join("downloads", chat_title, topic_title)
            os.makedirs(save_path, exist_ok=True)

            logger.info(f"Fetching photos from topic ID {topic_id}...")
            messages = client.iter_messages(
                chat_id,
                filter=InputMessagesFilterPhotos(),
                limit=None,
                reply_to=topic_id
            )
            messages_list = list(messages)
            total_photos = len(messages_list)

            if total_photos == 0:
                logger.info(f"No photos found in topic ID {topic_id}.")
                continue

            logger.info(f"Found {total_photos} photos in topic '{topic_title}'. Starting download...")
            found_photos = 0

            for message in tqdm(messages_list, total=total_photos, desc=f"Downloading from {topic_title}",
                                unit="photo"):
                if message.photo:
                    file_path = os.path.join(save_path, f"{message.id}.jpg")
                    if file_path not in downloaded_photos:
                        try:
                            # Download photo media
                            client.download_media(message.media, file=file_path)
                            if os.path.exists(file_path):
                                downloaded_photos[file_path] = True
                                found_photos += 1
                                save_downloaded_photos(DOWNLOADED_PHOTOS_FILE, downloaded_photos)
                            else:
                                logger.error(f"File {file_path} not found after download.")
                        except Exception as e:
                            logger.error(f"Error downloading photo {message.id}: {e}")
                            logger.info("Retrying in 60 seconds...")
                            time.sleep(60)
                            try:
                                client.download_media(message.media, file=file_path)
                                if os.path.exists(file_path):
                                    downloaded_photos[file_path] = True
                                    found_photos += 1
                                    save_downloaded_photos(DOWNLOADED_PHOTOS_FILE, downloaded_photos)
                                else:
                                    logger.error(f"File {file_path} not found after retry.")
                            except Exception as retry_error:
                                logger.error(f"Retry failed for photo {message.id}: {retry_error}")

            logger.info(f"Downloaded {found_photos} out of {total_photos} photos from topic '{topic_title}'.")
            # Save progress after processing each topic
            save_downloaded_photos(DOWNLOADED_PHOTOS_FILE, downloaded_photos)

    except Exception as e:
        logger.error(f"Error during photo download: {e}")
    finally:
        client.disconnect()


if __name__ == "__main__":
    # Load configuration and downloaded photos info
    config = load_config(CONFIG_FILE)
    api_id = config.get('api_id')
    api_hash = config.get('api_hash')
    phone_number = config.get('phone_number')
    chat_id = int(config.get('chat_id', 0))
    topics = config.get('topics', [])

    if not all([api_id, api_hash, phone_number, chat_id]):
        logger.error("Missing required configuration parameters.")
        exit(1)

    downloaded_photos = load_downloaded_photos(DOWNLOADED_PHOTOS_FILE)

    # Initialize Telegram client with a session name
    client = TelegramClient("session", api_id, api_hash)
    download_photos(client, phone_number, chat_id, topics, downloaded_photos)
