import schedule
import time
from neighborhoods_processing import process_neighborhoods_data
from products_processing import weekly_data_processing

def main():
    schedule.every().monday.at("10:00").do(process_neighborhoods_data, file_path="./dados/bairros.xlsx")

    schedule.every().monday.at("10:00").do(weekly_data_processing)

    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    main()
