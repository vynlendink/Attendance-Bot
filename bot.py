import discord
import os
import gspread
import json
from google.oauth2.service_account import Credentials

from discord.ext import commands
from dotenv import load_dotenv
from datetime import datetime
import random

intents = discord.Intents.all()
client = discord.Client(intents=intents)
bot = commands.Bot(command_prefix='.', intents=intents, case_insensitive=True)

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Use dynamic path for credentials
script_directory = os.path.dirname(os.path.abspath(__file__))
creds_path = os.path.join(script_directory, 'creds.json')
creds = Credentials.from_service_account_file(creds_path, scopes=scope)

sheets = gspread.authorize(creds)

spreadsheet_id = "1330-2-P2etVNhNd0WPYdvMAv_YoKVko0ya0vJ2rIv8o"
spreadsheet = sheets.open_by_key(spreadsheet_id)

classic_sheet = spreadsheet.worksheet("Classic")
retail_sheet = spreadsheet.worksheet("Live")
name_sheet = spreadsheet.worksheet("Names")

json_file_path = os.path.join(script_directory, 'channels.json')

@bot.command(aliases=['cata', 'c'])
async def classic(ctx, start_date: str, *, reason: str):
    await handle_date_range(ctx, classic_sheet, start_date, reason)

@classic.error
async def classic_error(ctx, error):
    if isinstance(error, commands.MissingRequiredArgument):
        await ctx.send("Please provide a start date (MM/DD/YY or MM/DD/YY-MM/DD/YY) and a reason. Example: .classic 07/04/24 New game release or .classic 07/04/24-07/10/24 Going on vacation.")
    elif isinstance(error, commands.BadArgument):
        await ctx.send("Please provide valid arguments. The dates should be in MM/DD/YY format.")
    else:
        await ctx.send("An error occurred while processing the command.")

@bot.command(aliases=['retail', 'l','r', 'tww', 'df'])
async def live(ctx, start_date: str, *, reason: str):
    await handle_date_range(ctx, retail_sheet, start_date, reason)

@live.error
async def live_error(ctx, error):
    if isinstance(error, commands.MissingRequiredArgument):
        await ctx.send("Please provide a start date (MM/DD/YY or MM/DD/YY-MM/DD/YY) and a reason. Example: .live 07/04/24 New game release or .live 07/04/24-07/10/24 Going on vacation.")
    elif isinstance(error, commands.BadArgument):
        await ctx.send("Please provide valid arguments. The dates should be in MM/DD/YY format.")
    else:
        await ctx.send("An error occurred while processing the command.")

def generate_unique_id(sheet):
    existing_ids = [row[0] for row in sheet.get_all_values()]
    while True:
        unique_id = ''.join(random.choices('0123456789', k=5))
        if unique_id not in existing_ids:
            return unique_id

async def handle_date_range(ctx, sheet, start_date, reason):
    current_date = datetime.now()
    formatted_now = current_date.strftime("%m/%d/%y")
    date2 = datetime.strptime(formatted_now, "%m/%d/%y")

    # Determine next available row
    nextAvailable = 1
    rowsToDelete = []
    for row in range(1, 50):
        if sheet.cell(row, 1).value:
            try:
                date1 = datetime.strptime(sheet.cell(row, 3).value, "%m/%d/%y")
                if date1 < date2:
                    rowsToDelete.append(row)
            except ValueError:
                continue
        if not sheet.cell(row, 1).value:
            nextAvailable = row
            break

    # Check if the user has an assigned name
    username = ctx.author.name
    for user in range(1, name_sheet.row_count):
        if name_sheet.cell(user, 1).value == username:
            username = name_sheet.cell(user, 2).value
            break
    

    # Process start_date to determine start and end dates
    dates = start_date.split('-')
    try:
        start_event_date = datetime.strptime(dates[0].strip(), "%m/%d/%y")
        if len(dates) == 2:
            end_event_date = datetime.strptime(dates[1].strip(), "%m/%d/%y")
            if end_event_date < start_event_date:
                await ctx.send("End date cannot be before the start date.")
                return
        else:
            end_event_date = start_event_date
    except ValueError:
        await ctx.send("Please provide valid dates in MM/DD/YY format.")
        return

    if start_event_date < date2:
        await ctx.send("You cannot enter a date that has already passed.")
        return

    # Format the dates for the spreadsheet
    formatted_start_date = start_event_date.strftime("%m/%d/%y")
    formatted_end_date = end_event_date.strftime("%m/%d/%y")
    unique_id = generate_unique_id(sheet)

    # Add data to the sheet
    data = [username, formatted_start_date, formatted_end_date, reason, unique_id]
    sheet.append_row(data)
    row_index = nextAvailable  # Use the next available row index

    # Overwrite to ensure proper date formatting
    sheet.update_cell(row_index, 2, formatted_start_date)
    sheet.update_cell(row_index, 3, formatted_end_date)

    await ctx.send(f"Posted from {formatted_start_date} to {formatted_end_date} with the reason: `{reason}` - ID: `{unique_id}`")

    # Insert an empty row after the data
    last_row_index = len(sheet.get_all_values()) + 1
    empty_row = [''] * sheet.col_count
    sheet.insert_rows([empty_row], row=last_row_index)

    if len(rowsToDelete) != 0:
        rowsToDelete.reverse()
        for row_index in rowsToDelete:
            sheet.delete_rows(row_index)
    
@bot.command(aliases=['nick','nickname','nn'])
async def name(ctx, username):
    not_found = False
    for user in range(1, name_sheet.row_count):
        if name_sheet.cell(user, 1).value == ctx.author.name:
            name_sheet.update_cell(user, 2, username)
            await ctx.send(f"Updated your username in the bot to {username}.")
            break
        elif name_sheet.cell(user, 1).value == None:
            not_found = True
            last_row_index = len(name_sheet.get_all_values()) + 1
            empty_row = [''] * name_sheet.col_count
            name_sheet.insert_rows([empty_row], row=last_row_index)
            break

    if not_found:
        data = [ctx.author.name, username]
        name_sheet.append_row(data)
        await ctx.send(f"Added you to the bot as {username}.")


@name.error
async def name_error(ctx, error):
    if isinstance(error, commands.MissingRequiredArgument):
        await ctx.send("You must provide a username.")
    else:
        await ctx.send("An error occurred while processing the command.")

@bot.command(aliases=['remove'])
async def remove_absence(ctx, absence_id: str):
    removed = False
    for user in range(1, name_sheet.row_count):
        if name_sheet.cell(user, 1).value == ctx.author.name:
            username = name_sheet.cell(user, 2).value.lower()
            break
        else:
            username = ctx.author.name
    sheets = [classic_sheet, retail_sheet]
    
    for sheet in sheets:
        for row in range(1, sheet.row_count + 1):  # Adjust the range as needed
            cell_value = sheet.cell(row, 5).value
            if cell_value == absence_id:
                if sheet.cell(row, 1).value.lower() == username:
                    sheet.delete_rows(row)
                    await ctx.send(f"Removed your absence with ID `{absence_id}`.")
                    removed = True
                    break
                else:
                    await ctx.send(f"You can only remove absences you have posted.")
                    removed = True
                    break
        if removed:
            break
    
    if not removed:
        await ctx.send(f"No absence found with ID `{absence_id}`.")

@bot.command(aliases=['?'])
async def h(ctx):
    await ctx.send("## DMG Attendance Bot Help Commands\n**.live, .retail, .l, .r, .tww, .df** - Post out on Live with the following arguments:\n   Date: MM/DD/YY format. Date ranges supported by using MM/DD/YY-MM/DD/YY\n   Reason: The reason you'll be absent, you don't need to go into too much detail, just so we have an idea of why.\n      Examples:\n         .live 07/24/24 Family coming over.\n         .live 07/24/24-07/31/24 Going to visit family.\n\n**.classic, .cata, .c** - Post out on Classic with the following arguments:\n   Date: MM/DD/YY format. Date ranges supported by using MM/DD/YY-MM/DD/YY\n   Reason: The reason you'll be absent, you don't need to go into too much detail, just so we have an idea of why.\n      Examples:\n         .classic 07/24/24 Family coming over.\n         .classic 07/24/24-07/31/24 Going to visit family.\n\n**.remove** - This will remove an absence with the ID provided when you posted it\n   Examples:\n      .remove 10234\n   Only the user who made the post is able to delete an absence.")

# Function to load channels for a guild from JSON
def load_channels(guild_id):
    if not os.path.exists(json_file_path):
        with open(json_file_path, 'w') as file:
            json.dump({}, file)
        return []

    with open(json_file_path, 'r') as file:
        data = json.load(file)

    return data.get(str(guild_id), [])

# Function to save channels for a guild to JSON
def save_channels(guild_id, channels):
    if not os.path.exists(json_file_path):
        data = {}
    else:
        with open(json_file_path, 'r') as file:
            data = json.load(file)

    data[str(guild_id)] = channels

    with open(json_file_path, 'w') as file:
        json.dump(data, file)

# Command for administrators to set authorized channels
@bot.command(aliases=['addchannel', 'ac'])
@commands.has_permissions(administrator=True)
async def set_channel(ctx, channel: discord.TextChannel):
    guild_id = ctx.guild.id
    channels = load_channels(guild_id)
    
    if channel.id not in channels:
        channels.append(channel.id)
        save_channels(guild_id, channels)
        await ctx.send(f"Channel {channel.mention} has been added to the authorized list for this guild.")
    else:
        await ctx.send(f"Channel {channel.mention} is already in the authorized list for this guild.")

@bot.command(aliases=['removechannel', 'rc'])
@commands.has_permissions(administrator=True)
async def remove_channel(ctx, channel: discord.TextChannel):
    guild_id = ctx.guild.id
    channels = load_channels(guild_id)
    
    if channel.id in channels:
        channels.remove(guild_id, channel.id)
        save_channels(guild_id, channels)
        await ctx.send(f"Channel {channel.mention} has been removed from the authorized list for this guild.")
    else:
        await ctx.send(f"Channel {channel.mention} is not in the authorized list for this guild.")

# Event to check channel before responding
@bot.event
async def on_message(message):
    # Ignore messages from the bot itself
    if message.author == bot.user:
        return
    
    # Ignore direct messages
    if message.guild is None:
        return
    
    guild_id = message.guild.id
    channels = load_channels(guild_id)
    

    if len(channels) == 0 or message.channel.id in channels:
        # Check if the message starts with the bot's command prefix
        if message.content.startswith(bot.command_prefix):
            ctx = await bot.get_context(message)
            if ctx.valid:
                await bot.process_commands(message)
        else:
            # Optionally, you can add a response here if the message is not a command
            # await message.channel.send("This message does not contain a valid command.")
            pass
    else:
        return



# Retrieve token from the .env file
load_dotenv()
bot.run(os.getenv('TOKEN'))