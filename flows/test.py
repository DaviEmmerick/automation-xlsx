import schedule
import time

def task():
  print("Aprovado!")

schedule.every().second.do(task)

while True:
  schedule.run_pending()
  time.sleep(1)