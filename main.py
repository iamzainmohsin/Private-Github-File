from database import ExcelHandler, Task
from reminder import Reminder
from config import TOKEN, CHANNEL_ID_general, CHANNEL_ID_notifications
# from keep_alive import keep_alive
from functools import wraps
import discord
from discord.ext import commands

#FILE NAME FOR DATABASE
# keep_alive()
file_name = "database.xlsx"


# Decorator: Restricting commands to the "general" channel
def only_general_channel():
    def decorator(func):
        @wraps(func)
        async def wrapper(ctx, *args, **kwargs):
            if ctx.channel.name != "task_manager":
                await ctx.send("This command can only be used in the `general` channel.")
                return
            return await func(ctx, *args, **kwargs)
        return wrapper
    return decorator


# Initialize bot and database
bot = commands.Bot(command_prefix="!", intents=discord.Intents.all())
tasks = ExcelHandler(file_name)
reminder = Reminder(tasks, bot, CHANNEL_ID_notifications)
tasks.workbook_setup()

    
# Bot event handlers
@bot.event
async def on_ready():
    print(f"Logged in as {bot.user}")
    general_channel = bot.get_channel(CHANNEL_ID_general) # ---> Replace
    await general_channel.send("✅ BOT is now active!")
    notification_channel = bot.get_channel(CHANNEL_ID_notifications) # ---> Replace
    if notification_channel:
        await notification_channel.send("✅ Reminder system is now active!")

    bot.loop.create_task(reminder.refresh_tasks(300))  # Refresh every 5 minutes
    bot.loop.create_task(reminder.schedule_all_reminders())  # Continuously check for new reminders
    print("Reminder system has started!")


#Handles Relevant Commands:
def bot_commands():
    #Adds Task to the Sheet!
    @bot.command(name="create")
    @only_general_channel()
    async def add_task(ctx, *, input_data: str):
        try:
            description, due_date, status = map(str.strip, input_data.split(","))
            new_task = Task(description, due_date, status)
            result = tasks.add_tasks(new_task)
            if result:
                await ctx.send(f"Task '{description}' added successfully!")
            else:
                await ctx.send("Failed to add the task. Please check your input!")
        except ValueError:
            await ctx.send("Invalid input format! Use: !add_task <description>, <due_date>, <status>")


    #View Tasks In the Sheet!
    @bot.command(name="show")
    @only_general_channel()
    async def view_tasks(ctx):
        all_tasks = tasks.get_tasks()
        if all_tasks:
            formatted_tasks = "\n".join([
            f"**{i+1}.** **Task:** '{task.description}'\n   **Due Date:** '{task.due_date}'\n   **Status:** {task.status}\n"
            for i, task in enumerate(all_tasks)
        ])
            # Send the formatted output to Discord
            await ctx.send(f"\n**ACTIVE TASKS**\n\n{formatted_tasks}")
        else:
            await ctx.send("No tasks found in the Excel sheet.")    


    # Marking the Tasks Complete
    @bot.command(name="complete")
    @only_general_channel()
    async def mark_complete(ctx, task_index: int):
        try:
            tasks.complete_task(task_index, "C")
            tasks.move_complete_task()
            await ctx.send(f"Task {task_index} marked as complete and moved to 'Completed'.")
        except Exception as e:
            await ctx.send(f"Error marking task {task_index} as complete: {str(e)}")


    #Deleting the Tasks:
    @bot.command(name="remove")
    @only_general_channel()
    async def delete_task(ctx, task_id: str):
        if not task_id.isdigit():
            await ctx.send(f"Invalid Task ID: '{task_id}'. Please use numeric values only.")
            return

        # Call the delete_task method
        try:
            task_deleted = tasks.delete_task(task_id)
            if task_deleted:
                await ctx.send(f"Task ID '{task_id}' successfully deleted")
            else:
                await ctx.send(f"Task ID '{task_id}' not found in the sheet.")
        except Exception as e:
            await ctx.send(f"An error occurred while deleting the task: {str(e)}")


    #Shows A Summary of the Progress
    @bot.command(name="progress")
    @only_general_channel()
    async def summary(ctx):
        try:
            report_summary = tasks.generate_report()
            await ctx.send(f"\n**SUMMARY**\n{report_summary}")
        except Exception as e:
            await ctx.send(f"Error generating report: {str(e)}")


def main():
    bot_commands()
    bot.run(TOKEN)
    
    
if __name__ == '__main__':
    main()       





