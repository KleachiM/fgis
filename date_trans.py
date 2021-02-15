import datetime, sys

DAYS_DIFF = 1  # если 1 то сортировка за все даты до вчерашнего дня, если 0 то до сегодня
DAYS_COUNT = 26


def get_date(days):
    date_from = datetime.date.today() - datetime.timedelta(days=days)

    if datetime.datetime.utcnow().year - date_from.year == 1:  # если 40 рабочих дней назад был прошлый год
        date_to = datetime.date(date_from.year, 12, 31)
        date_begin_new_year = datetime.date(datetime.datetime.utcnow().year, 1, 1)
        date_end_new_year = datetime.date.today() - datetime.timedelta(days=DAYS_DIFF)
        verification_old_year = date_from.year
        verification_new_year = date_end_new_year.year

        return {
            'past_year_from': date_from.isoformat(),
            'past_year_to': date_to.isoformat(),
            'past_verification_year': verification_old_year,
            'this_year_from': date_begin_new_year.isoformat(),
            'this_year_to': date_end_new_year.isoformat(),
            'this_verification_year': verification_new_year
        }

    elif datetime.datetime.utcnow().year == date_from.year:  # если 40 рабочих дней назад был текущий год
        date_to = datetime.date.today() - datetime.timedelta(days=DAYS_DIFF)
        verification_year = date_from.year

        return {
            'date_from': date_from.isoformat(),
            'date_to': date_to.isoformat(),
            'verification_year': verification_year
        }
    else:
        print('Необходима проверка дат')
        sys.exit(0)

