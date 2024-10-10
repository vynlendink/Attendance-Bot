import discord
import os
import gspread
import json
import re
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
professions_sheet = spreadsheet.worksheet("Professions")

json_file_path = os.path.join(script_directory, 'channels.json')

@bot.command(aliases=['profession', 'proff', 'p'])
async def add_profession(ctx, *, profession):
    discord_tag = str(ctx.author)
    profession = profession.lower()
    professions_sheet.append_row([discord_tag, profession])
    await ctx.send(f'{discord_tag} registered profession: {profession}')

@bot.command(name='craft')
async def request_craft(ctx, *, item):
    item = item.lower()

    if ctx.author.id == 163780323296018432:
        await ctx.author.send("kig do NOT cum in the professions channel")

    if "cum" in item:
        await ctx.send("kig you have got to be normal")
        return
    
    item_words = item.split()

    if len(item_words) == 1 and item.endswith('s'):
        item_pattern = re.compile(rf'\b{item[:-1]}(s)?\b')
    elif item_words[-1].endswith('s') and len(item_words) >= 2:
        last_word = re.escape(item_words[-1][:-1])
        item_pattern = re.compile(rf'^\b{" ".join(re.escape(word) for word in item_words[:-1])} {last_word}(s)?\b$', re.IGNORECASE)
    elif item_words[-1].endswith('s') and len(item_words) == 1:
        item_pattern = re.compile(rf'\b{item[:-1]}(s)?\b')
    else:
        item_pattern = re.compile(rf'\b{item}(s)?\b')

    all_data = professions_sheet.get_all_values()
    all_items = [(row[0], row[1].lower()) for row in all_data[1:] if len(row) >= 2]

    matching_items = [row for row in all_items if item_pattern.search(row[1])]
    
    if not matching_items:
        await ctx.send(f'No users found with the profession: {item}')
        return
    
    user_ids = set()
    
    for user_name, matching_item in matching_items:
        user = ctx.guild.get_member_named(user_name)
        if user and user.id not in user_ids:
            user_ids.add(user.id)
    
    user_mentions = ' '.join([f'<@{user_id}>' for user_id in user_ids])
    
    if user_mentions:
        if len(user_ids) == 1:
            message = await ctx.send(f'{user_mentions}, {ctx.author.display_name} needs you to craft {item}')
        elif len(user_ids) == 2:
            message = await ctx.send(f'{user_mentions}, {ctx.author.display_name} needs either of you to craft {item}')
        else:
            message = await ctx.send(f'{user_mentions}, {ctx.author.display_name} needs one of you to craft {item}')

        await message.add_reaction("✅")

        def check(reaction, user):
            return (
                (user.id in user_ids or user.id == ctx.author.id) and 
                str(reaction.emoji) == "✅" and 
                reaction.message.id == message.id
            )

        reaction, user = await bot.wait_for('reaction_add', check=check)
        await message.delete()
        await ctx.message.delete()
        await ctx.author.send(f'<@{user.id}> has finished your crafting order')
    else:
        await ctx.send(f'No users found with the profession: {item}')


@bot.command(aliases=['removeproff', 'removep'])
async def remove_profession(ctx, *, profession=None):

    if not profession:
        await ctx.send("You must input a profession to remove")
        return
    
    profession = profession.lower()
    discord_tag = str(ctx.author)

    profession_data = professions_sheet.get_all_values()

    professions = [(row[0], row[1].lower()) for row in profession_data if row[0]]

    rows_to_delete = []
    for i, (user, registered_profession) in enumerate(professions, start=1):
        if user == discord_tag and registered_profession == profession:
            rows_to_delete.append(i)

    if rows_to_delete:
        for row in sorted(rows_to_delete, reverse=True):
            professions_sheet.delete_rows(row)
        
        await ctx.send(f'Removed profession `{profession}` for {ctx.author.display_name}')
    else:
        await ctx.send(f'No profession found `{profession}` for user {ctx.author.display_name}')


@bot.command(name='registered')
async def registered(ctx):
    data = professions_sheet.get_all_values()
    
    professions_dict = {}
    
    for row in data:
        crafter = row[0]
        profession = row[1].lower().title()
        
        if profession in professions_dict:
            professions_dict[profession].append(crafter)
        else:
            professions_dict[profession] = [crafter]
    
    sorted_professions = sorted(professions_dict.items())
    
    message_lines = []
    current_message = ''
    
    for profession, crafters in sorted_professions:
        mentions = ', '.join([f'<@{ctx.guild.get_member_named(crafter).id}>' for crafter in crafters if ctx.guild.get_member_named(crafter)])
        line = f'{profession} can be crafted by {mentions}\n'
        
        if len(current_message) + len(line) > 2000:
            message_lines.append(current_message)
            current_message = line
        else:
            current_message += line

    if current_message:
        message_lines.append(current_message)

    for message in message_lines:
        await ctx.author.send(message)


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
async def live(ctx, start_date: str, *, reason="No reason provided"):
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
        unique_id = ''.join(random.choices('123456789', k=5))
        if unique_id not in existing_ids:
            return unique_id

async def handle_date_range(ctx, sheet, start_date, reason):
    current_date = datetime.now()
    formatted_now = current_date.strftime("%m/%d/%y")
    date2 = datetime.strptime(formatted_now, "%m/%d/%y")

    all_values = sheet.get_all_values()
    rowsToDelete = []
    nextAvailable = len(all_values) + 1

    for row, values in enumerate(all_values, start=1):
        if values[0]:
            try:
                date1 = datetime.strptime(values[2], "%m/%d/%y")
                if date1 < date2:
                    rowsToDelete.append(row)
            except ValueError:
                continue
        else:
            nextAvailable = row
        break

    username = ctx.author.name
    for user in range(1, name_sheet.row_count):
        if name_sheet.cell(user, 1).value == username:
            username = name_sheet.cell(user, 2).value
            break

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

    formatted_start_date = start_event_date.strftime("%m/%d/%y")
    formatted_end_date = end_event_date.strftime("%m/%d/%y")
    unique_id = generate_unique_id(sheet)

    data = [username, formatted_start_date, formatted_end_date, reason, unique_id]
    sheet.append_row(data)
    row_index = nextAvailable

    sheet.update_cell(row_index, 2, formatted_start_date)
    sheet.update_cell(row_index, 3, formatted_end_date)

    await ctx.send(f"Posted from {formatted_start_date} to {formatted_end_date} with the reason: `{reason}` - ID: `{unique_id}`")

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
    sheets = [retail_sheet, classic_sheet]
    
    for sheet in sheets:
        for row in range(1, sheet.row_count + 1):
            cell_value = sheet.cell(row, 5).value
            if not cell_value:
                break
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

@bot.command(aliases=['edit'])
async def edit_absence(ctx, absence_id: str, start_date: str = None, *, reason: str = None):
    username = ctx.author.name.lower()
    name_rows = name_sheet.get_all_values()
    
    for row in name_rows:
        if row[0].lower() == username:
            username = row[1].lower()
            break

    sheets = [retail_sheet, classic_sheet]
    updated = False

    for sheet in sheets:
        rows = sheet.get_all_values() 
        for row_idx, row in enumerate(rows):
            cell_value = row[4]
            if cell_value == absence_id and row[0].lower() == username:
                if start_date:
                    try:
                        dates = start_date.split('-')
                        start_event_date = datetime.strptime(dates[0].strip(), "%m/%d/%y")
                        end_event_date = start_event_date
                        if len(dates) == 2:
                            end_event_date = datetime.strptime(dates[1].strip(), "%m/%d/%y")
                        row[1] = start_event_date.strftime("%m/%d/%y").lstrip("'")
                        row[2] = end_event_date.strftime("%m/%d/%y").lstrip("'")
                    except ValueError:
                        await ctx.send("Please provide valid dates in MM/DD/YY format.")
                        return

                if reason:
                    row[3] = reason 

                sheet.update(range_name=f'A{row_idx + 1}:E{row_idx + 1}', values=[row], value_input_option='USER_ENTERED')
                
                await ctx.send(f"Updated absence with ID `{absence_id}`.")
                updated = True
                break

        if updated:
            break

    if not updated:
        await ctx.send(f"No absence found with ID `{absence_id}`.")


@bot.command(aliases=['?'])
async def h(ctx):
    await ctx.send("""## DMG Attendance Bot Help Commands
                   **.live, .retail, .l, .r, .tww, .df** - Post out on Live with the following arguments:
                      Date: MM/DD/YY format. Date ranges supported by using MM/DD/YY-MM/DD/YY
                      Reason: The reason you'll be absent, you don't need to go into too much detail, just so we have an idea of why.
                         Examples:
                            .live 07/24/24 Family coming over.
                            .live 07/24/24-07/31/24 Going to visit family.
                   
                   **.profession, .proff, .p** - Set what items you're capable of crafting
                       Examples: .proff warglaive

                   **.craft** - Request a craft for an item
                       Examples: .craft cloth bracer
                   
                   **.removeproff, .removep** - Removes all of your professions, mostly so when the tier is over we can make Tainted do his job.

                   **.classic, .cata, .c** - Post out on Classic with the following arguments:
                      Date: MM/DD/YY format. Date ranges supported by using MM/DD/YY-MM/DD/YY
                      Reason: The reason you'll be absent, you don't need to go into too much detail, just so we have an idea of why.
                         Examples:
                            .classic 07/24/24 Family coming over.
                            .classic 07/24/24-07/31/24 Going to visit family.
                   
                   **.remove** - Remove an absence with the ID provided when you posted it
                      Examples:
                         .remove 10234
                      Only the user who made the post is able to delete an absence.
                      
                   **.edit** Edit an absence with the ID provided when you posted it
                      Examples:
                         .edit 10234 07/24/24 Family leaving at 8, will be an hour late.
                      Only the user who made the post is able to edit an absence""")

def load_channels(guild_id):
    if not os.path.exists(json_file_path):
        with open(json_file_path, 'w') as file:
            json.dump({}, file)
        return []

    with open(json_file_path, 'r') as file:
        data = json.load(file)

    return data.get(str(guild_id), [])

def save_channels(guild_id, channels):
    if not os.path.exists(json_file_path):
        data = {}
    else:
        with open(json_file_path, 'r') as file:
            data = json.load(file)

    data[str(guild_id)] = channels

    with open(json_file_path, 'w') as file:
        json.dump(data, file)

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


@bot.event
async def on_message(message):
    if message.author == bot.user:
        return
    
    if message.guild is None:
        return
    
    guild_id = message.guild.id
    channels = load_channels(guild_id)
    

    if len(channels) == 0 or message.channel.id in channels:
        if message.content.startswith(bot.command_prefix):
            ctx = await bot.get_context(message)
            if ctx.valid:
                await bot.process_commands(message)
        else:
            pass
    else:
        return



# Retrieve token from the .env file
load_dotenv()
bot.run(os.getenv('TOKEN'))
