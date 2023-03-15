from os.path import join
from re import sub
from shutil import move
from zipfile import BadZipFile

from scripts.check_file import check_file
from scripts.find import find
from scripts.save_excel import save_excel
from time import time
import logging
from os import mkdir, path, listdir


logger = logging.getLogger()
logging.basicConfig(filename=f'logs.log',
                    filemode='a',
                    format='%(levelname)s: %(asctime)s %(message)s',
                    datefmt='%d.%m.%Y %H:%M:%S',
                    level=logging.DEBUG)


def move_file(file: str, dir_move: str):
    zip_file = sub('.gsfx', '.zip', file)
    xml_file = sub('.gsfx', '.xml', file)

    if not path.exists(dir_move):
        mkdir(dir_move)
    if path.exists(join('download', file)):
        move(join('download', file), join(dir_move, file))
    elif path.exists(join('download', zip_file)):
        move(join('download', zip_file), join(dir_move, zip_file))
    elif path.exists(join('download', file)):
        move(join('download', xml_file), join(dir_move, xml_file))


def main() -> None:

    if not path.exists('download'):
        mkdir('download')
        logger.info('Пустая папка')
        return

    files = listdir('download')
    if '.gitkeep' in files:
        files.remove('.gitkeep')

    if files is None:
        logger.info('Пустая папка')
        return

    for file in files:

        try:
            logger.debug(f'START {file}')
            name = check_file(file)
            info_smeta = find(name)
            name_form = 'pattern_VOR.xlsx'
            save_excel(name, name_form, info_smeta)
            logger.debug(f'STOP GOOD {file}')

        except FileNotFoundError as error:
            logger.error(f'FileNotFoundError Не найден {error.filename}')
            logger.debug(f'STOP BAD {file}')

            continue

        except TypeError as error:
            logger.error(f'TypeError {error.args} {error}  {file}')
            logger.debug(f'STOP BAD {file}')
            move_file(file, 'bad_files')
            continue

        except BadZipFile as error:

            logger.error(f'BadZipFile {error.args} {error} {file}')
            logger.debug(f'STOP BAD {file}')

            continue

        except KeyError as error:
            logger.error(f' KeyError {error.args}  {error}  {file}')
            logger.debug(f'STOP BAD {file}')
            move_file(file, 'bad_files')
            continue


if __name__ == '__main__':

    start = time()

    main()

    stop = time()
    t = stop - start

    logger.info(f'Время работы: {round(t, 2)}с')
