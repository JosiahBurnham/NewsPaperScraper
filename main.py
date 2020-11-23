import sys
import schedule
import time
import pathlib

sys.path.insert(1, str(pathlib.Path(__file__).parent.absolute()) + "\\resources")
import Email_Paper
import Make_Paper


def create_newspaper(t):
    """ creates the newspaper

    creates the newspaper and sends the email to
    whoever is specified

    :param t:
     this is just required for the schedule module
    :return:
    """
    Make_Paper.MakePaper().make_paper()

    Email_Paper.EmailHelper().send_email("RECIPENT"
                                         , "YOUR SUBJECT"
                                         , "Here is today's copy of your favorite newspaper."
                                         , "YOUR PATH")
    print("I SENT AN EMAIL", t)


schedule.every().day.at("19:42").do(create_newspaper, 'It is 19:42')

while True:
    schedule.run_pending()
    time.sleep(60)  # wait one minute
