import os
import random

import sys
import time
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from models import VkClientProxy
from utils import logger

load_dotenv()

NUM_USERS_BEFORE_ASK = int(os.getenv('NUM_USERS_BEFORE_ASK'))
MIN_WAIT = int(os.getenv('MIN_WAIT'))
MAX_WAIT = int(os.getenv('MAX_WAIT'))


def process_sheet(client, sheet, invite_msg, num_friends):
    row_count = min(num_friends, sheet.max_row)
    success_rows = 0

    for idx, r in enumerate(range(2, row_count + 1), 1):
        try:
            id = sheet[f'D{str(r)}'].value

            # vk_method = VkApiMethod(client._obj)
            resp = client._obj._vk.method('friends.add', {'user_id': int(id), 'text': invite_msg})
            # resp = client.friends.add(int(id))
            # resp = vk_method.friends.add(int(id))

            logger.info(f'Added friend with id={id} with result {resp}')

            success_rows += 1

            wait_slot = random.randint(MIN_WAIT, MAX_WAIT)
            logger.info(f'Sleeping for {wait_slot}s...')
            time.sleep(wait_slot)

        except Exception as ex:
            logger.error(f'Failed to add friend {id}: {ex}')

        finally:
            if (idx % NUM_USERS_BEFORE_ASK) == 0:
                print('\n---------------------------------------------------------')
                print('EXCEEDED MAX USER LIMIT')
                input('PLEASE, PRESS [ENTER] TO CONTINUE...')

    return f'{success_rows}/{row_count}'


def process_file(client, filename):
    wb = load_workbook(filename)
    ws: Worksheet = wb.active
    sheet = wb.worksheets[0]
    sheet_map = {sheet.title: sheet for sheet in wb.worksheets}
    row_count = sheet.max_row
    prev_login = None
    for r in range(2, row_count + 1):
        try:
            login = ws[f'B{str(r)}'].value
            sheetname = ws[f'C{str(r)}'].value
            invite_msg = ws[f'D{str(r)}'].value
            num_friends = ws[f'E{str(r)}'].value

            sheet = sheet_map.get(sheetname)
            if not sheet:
                raise RuntimeError(f'Sheet with name {sheetname} is not found')
            if not prev_login or prev_login != login:
                if login not in [account[0] for account in client._accounts]:
                    raise RuntimeError(f'Account with login {sheetname} is not found')
                else:
                    client.auth_as(login)
                prev_login = login

            logger.info(f'Starting {login} for audience {sheetname}...')
            resp = process_sheet(client, sheet, invite_msg, num_friends)
            logger.info(f'Finished {login} for audience {sheetname}')
            # time.sleep(random.randint(MIN_WAIT, MAX_WAIT))
        except Exception as ex:
            logger.error(f'Failed to send message for account {login}: {ex}')

    wb.save(filename)


def main():
    vk_client = VkClientProxy()
    vk_client.load_accounts()
    # vk_client.auth()

    if len(sys.argv) > 1:
        param = sys.argv[1]
        process_file(vk_client, param)


if __name__ == '__main__':
    main()
