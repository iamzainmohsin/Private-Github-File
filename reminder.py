import asyncio
from datetime import datetime, timedelta
import time
import random

class Reminder():
    def __init__(self, excel_handler, bot, channel_ID):
        self.scheduled_tasks = set()
        self.excel_handler = excel_handler
        self.bot = bot
        self.channel_ID = channel_ID 
        self.last_sent = {}


    async def refresh_tasks(self, interval):
        while True:
            print("Refreshing tasks...")
            current_tasks = self.fetch_pending_tasks()  # Fetch latest tasks from Excel

            # Always check and reschedule tasks
            for task in current_tasks:
                if task.index not in self.scheduled_tasks:
                    # Add new tasks to the scheduler
                    print(f"Scheduling new task: {task.description}")
                    self.scheduled_tasks.add(task.index)
                    asyncio.create_task(self.schedule_reminder(task))
                else:
                    # Re-check frequency for already scheduled tasks
                    print(f"Rechecking task: {task.description}")
                    asyncio.create_task(self.schedule_reminder(task))  # Ensure reminders continue

            await asyncio.sleep(interval)  # Wait for the specified interval (e.g., 5 minutes)


    #Deletes up overdue tasks after 3 days
    def clean_overdue_tasks(self):
        print("Cleaning overdue tasks...")
        current_tasks = self.fetch_pending_tasks()
        today = datetime.now().date()
        removed_tasks = []

        print(f"Total pending tasks: {len(current_tasks)}")

        for task in current_tasks:
            try:
                due_date = datetime.strptime(task.due_date, "%d-%B-%Y").date()
            except ValueError:
                print(f"Skipping task '{task.description}': Invalid due date format -> {task.due_date}")
                continue

            days_overdue = (today - due_date).days
            print(f"Checking task: {task.description}, Due Date: {due_date}, Overdue by: {days_overdue} days")

            if days_overdue > 3:
                removed_tasks.append(task)
                self.excel_handler.delete_task(str(task.index))  # Ensure task ID is passed as string

        print(f"{len(removed_tasks)} overdue tasks removed.")


    #Loads the Pending tasks from the sheet
    def fetch_pending_tasks(self):
        # Directly call the method from the ExcelHandler class
        pending_tasks = self.excel_handler.get_tasks()
        print(f"Pending Tasks Fetched: {len(pending_tasks)} tasks found.")
        return pending_tasks
  

    def days_until_due(self, due_date):
        today = datetime.now().date()
        
        if isinstance(due_date, str):
            due_date = datetime.strptime(due_date, "%d-%B-%Y").date()
        
        days_left = (due_date - today).days
        return days_left


    #Sets Reminder Frequency
    def reminder_frequency(self, days_left):
        if days_left < 0:
            return "Immediate Warning"  # Overdue tasks
        elif days_left == 0:
            return "Hourly Reminder"  # Tasks due today
        elif days_left <= 3:
            return "Every 4 Hours"  # Tasks due within 3 days
        elif days_left <= 7:
            return "Daily Reminder"  # Tasks due later in the week
        else:
            return "Weekly Reminder"  # Far future tasks

          
    #Schedules the reminders
    async def schedule_reminder(self, task):
        task_description = task.description
        due_date = task.due_date
        days_left = self.days_until_due(due_date)
        frequency = self.reminder_frequency(days_left)
        channel = self.bot.get_channel(self.channel_ID)

        if not channel:
            print("Channel not found!")
            return

        # Intervals for different frequencies (in seconds)
        intervals = {
        "Immediate Warning": 0,
        "Hourly Reminder": 3600,       # 1 hour -->3600
        "Every 4 Hours": 14400,       # 4 hours --> 14400
        "Daily Reminder": 86400,      # 24 hours --> 86400
        "Weekly Reminder": 604800     # 7 days --> 604800
    }
        interval = intervals.get(frequency, None)

        initial_delay = random.randint(5, 30)  # Random delay between 5 to 30 seconds
        print(f"Task '{task_description}' will start reminders after {initial_delay} seconds.")
        await asyncio.sleep(initial_delay)

        # Track last sent time using `last_sent`
        if task.index not in self.last_sent:
            self.last_sent[task.index] = time.time()  # Initialize last sent time

        while True:
            # Check if the task is still active
            current_tasks = self.fetch_pending_tasks()
            if not any(t.index == task.index for t in current_tasks):
                print(f"Task '{task_description}' is no longer active. Stopping reminders.")
                break

            # Calculate time since the last reminder
            current_time = time.time()
            last_sent_time = self.last_sent[task.index]

            if current_time - last_sent_time >= interval:
                # Send the reminder
                message = f"<@774273998078214165>\n📌 **Task Reminder:** {task_description}\n**Due Date:** {due_date}\n**Reminder Frequency:** {frequency}"
                await channel.send(message)

                # Update the last sent time
                self.last_sent[task.index] = current_time
                print(f"Reminder sent for task: {task_description} at {time.ctime(current_time)}")
            else:
                # Debugging: Show why the task is waiting
                remaining_time = interval - (current_time - last_sent_time)
                print(f"Task '{task_description}' waiting for {remaining_time} seconds before next reminder.")

            # Sleep before rechecking
            await asyncio.sleep(10)


    #Looks for newly added tasks within the sheet
    async def check_and_update_tasks(self):
        while True:
            new_tasks = self.fetch_pending_tasks()

            for i, task in enumerate(new_tasks):  # Generate task_id dynamically
                task_id = i + 1  # Task ID based on enumerate()
                
                if task_id not in self.scheduled_tasks:
                    self.scheduled_tasks.add(task_id)
                    # Introduce staggered delay before scheduling
                    await asyncio.sleep(i * 3)
                    asyncio.create_task(self.schedule_reminder(task))
                        
            await asyncio.sleep(300)  # Check for new tasks every 5 minutes

            
    #Runs the clean_over_due() function
    async def daily_cleanup(self):
        while True:
            self.clean_overdue_tasks()
            await asyncio.sleep(86400)  # 24 hours


    #Runs Everything:
    async def schedule_all_reminders(self):
        asyncio.create_task(self.daily_cleanup())
        await self.check_and_update_tasks()


    
        
