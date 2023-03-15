import os
from os.path import join
from zipfile import ZipFile
from re import sub
from logging import getLogger


log = getLogger()


def check_file(name_smeta: str) -> str:
    """
    Проверка. Сметный ли это файл?

    :param name_smeta: Имя сметного файла
    :type name_smeta:str
    :return: Возвращает name
    :rtype: str
    :exception: KeyError, FileExistsError, TypeError
    """

    path = join(r'download')

    if name_smeta.endswith(r'.gsfx'):

        name = name_smeta[:-5]
        new_name = sub(r'.gsfx', r'.zip', name_smeta)
        os.rename(join(path, name_smeta), join(path, new_name))

        file_zip = join(path, new_name)
        lst = [f.filename for f in ZipFile(file_zip).infolist()]

        if 'Data.xml' not in lst:
            raise KeyError

        with ZipFile(file_zip, 'r') as file:
            file.extract('Data.xml', path)
        os.remove(join(path, new_name))
        new_name = f'{name}.xml'

        try:
            os.rename(join(path, 'Data.xml'), join(path, new_name))
        except FileExistsError as error:
            log.error(f'{error} {name}')
            name = f'_{name}_'
            new_name = f'{name}.xml'
            os.rename(join(path, 'Data.xml'), join(path, new_name))
            os.rename(join(path, 'Data.xml'), join(path, new_name))

        return name

    elif name_smeta.endswith(r'.xml'):
        name = name_smeta[:-4]
        return name

    else:

        raise TypeError
