#! /usr/bin/env python3

import sys
import csv
import logging
import logging.handlers

import click
from arrow.arrow import Arrow
from datetime import timedelta
from ics import Calendar
from ics.grammar.parse import ParseError


logger = logging.getLogger('ical2csv')

LOG_FILENAME = 'ical2csv.log'


def read_ical_file(filename):
    """Reads an iCalendar file and returns all events from that file."""
    logger.info('Reading iCalendar file...')
    with open(filename, 'r') as calendar_file:
        c = Calendar(calendar_file.read())
    logger.info('iCalendar file read.')
    return c.events

def write_csv_file(events, filename):
    attr = ['name', 'begin', 'end', 'duration', 'uid', 'description', 'created',
            'last_modified', 'location', 'url', 'transparent', 'alarms',
            'attendees', 'categories', 'status', 'organizer', 'classification']
    logger.info('Writing CSV file...')
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=attr)
        writer.writeheader()
        # iterate over all events and add data to table
        for event in events:
            writer.writerow( { k: marshal_data(getattr(event, k)) for k in attr } )
    logger.info('CSV file written.')

def marshal_data(data):
    """Marshals the data for export into a CSV file."""
    if not data:
        return ''
    elif type(data) == str:
        return data
    elif type(data) == Arrow:
        return data.format('DD.MM.YYYY HH:mm:ss')
    elif type(data) == timedelta:
        return str(data)
    elif type(data) == set:
        return ', '.join([x for x in data])
    else:
        logger.error('Unsupported type: {}'.format(type(data)))
        raise ValueError()

@click.command()
@click.option('--verbose', '-v', is_flag=True, help='Enables verbose mode.', default=False)
@click.argument('iCalendar_file', type=click.Path(exists=True)) # iCalendar file that should be converted into CSV format
@click.argument('CSV_file', type=click.Path(exists=False, writable=True), default=None, required=False) # CSV file that should be created. If no value was given, the file extension .csv will be added to the iCalendar file.
@click.version_option('0.1')
def convert_ical_file(verbose, icalendar_file, csv_file):
    """Simple tool to convert an iCalendar file into a CSV file to be used in LibreOffice Calc."""
    try:
        if csv_file == None:
            csv_file = '{}.csv'.format(icalendar_file)
            logger.info('Writing to CSV file: {}'.format(csv_file))
        write_csv_file(read_ical_file(icalendar_file), csv_file)
    except UnicodeDecodeError as e:
        logger.error('Error while reading file.')
        logger.error(e)
    except NotImplementedError as e:
        logger.error('iCalendar file contains multiple calendars. This is currently not supported.')
        logger.error(e)
    except ParseError as e:
        logger.error('iCalendar file not valid.')
        logger.error(e)

def create_logger():
    # create logger for this application
    global logger
    logger.setLevel(logging.DEBUG)
    log_to_file = logging.handlers.RotatingFileHandler(LOG_FILENAME, maxBytes=262144,
                                                       backupCount=5, encoding='utf-8')
    log_to_file.setLevel(logging.DEBUG)
    logger.addHandler(log_to_file)
    log_to_screen = logging.StreamHandler(sys.stdout)
    log_to_screen.setLevel(logging.INFO)
    logger.addHandler(log_to_screen)

if __name__ == '__main__':
    create_logger()
    convert_ical_file()
