import logging
import azure.functions as func
from Test import process_pptx_files 

app = func.FunctionApp()

@app.timer_trigger(schedule="0 10 * * * *", arg_name="myTimer", run_on_startup=False, use_monitor=False)
def testtimer(myTimer: func.TimerRequest) -> None:
    if myTimer.past_due:
        logging.info('The timer is past due!')

    logging.info('Timer trigger function executed.')

    try:

        process_pptx_files()
    except Exception as e:
        logging.error(f"Error executing pptx processing: {str(e)}")
