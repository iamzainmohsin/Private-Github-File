import random
import asyncio

# List of random messages
random_messages = [
    "Hello there!", "How's everyone doing?",
    "Don't forget to complete your tasks!", "Here's a random message! :)",
    "Keep up the great work!"
]


async def send_random_message(client, channel_id):
    # Access the channel by ID
    channel = client.get_channel(channel_id)

    while True:
        # Choose a random message
        message = random.choice(random_messages)

        # Send the random message to the channel
        await channel.send(message)

        # Wait for 1 minute (or adjust as needed)
        await asyncio.sleep(120)
