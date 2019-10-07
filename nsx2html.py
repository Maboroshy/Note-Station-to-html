#!/usr/bin/env python

import re
import sys
import time
import json
import zipfile
import collections

from pathlib import Path


# You can adjust some setting here. Default is for QOwnNotes app.
media_dir_name = 'media'  # name of the directory inside the produced directory where all images and attachments will be stored
creation_date_in_filename = False  # True to insert note creation time to the note file name, False to disable

############################################################################

Notebook = collections.namedtuple('Notebook', ['path', 'media_path'])


def sanitise_path_string(path_str):
    for char in (':', '/', '\\', '|'):
        path_str = path_str.replace(char, '-')
    for char in ('?', '*'):
        path_str = path_str.replace(char, '')
    path_str = path_str.replace('<', '(')
    path_str = path_str.replace('>', ')')
    path_str = path_str.replace('"', "'")

    return path_str[:240]


work_path = Path.cwd()
media_dir_name = sanitise_path_string(media_dir_name)


if len(sys.argv) > 1:
    files_to_convert = [Path(path) for path in sys.argv[1:]]
else:
    files_to_convert = Path(work_path).glob('*.nsx')

if not files_to_convert:
    print('No .nsx files found')
    exit(1)

for file in files_to_convert:
    nsx_file = zipfile.ZipFile(str(file))
    config_data = json.loads(nsx_file.read('config.json').decode('utf-8'))
    notebook_id_to_path_index = {}

    recycle_bin_path = work_path / Path('Recycle bin')

    n = 1
    while recycle_bin_path.is_dir():
        recycle_bin_path = work_path / Path('{}_{}'.format('Recycle bin', n))
        n += 1

    recycle_bin_media_path = recycle_bin_path / media_dir_name
    recycle_bin_media_path.mkdir(parents=True)
    notebook_id_to_path_index['1027_#00000000'] = Notebook(recycle_bin_path, recycle_bin_media_path)

    print('Extracting notes from "{}"'.format(file.name))

    for notebook_id in config_data['notebook']:
        notebook_data = json.loads(nsx_file.read(notebook_id).decode('utf-8'))
        notebook_title = notebook_data['title'] or 'Untitled'
        notebook_path = work_path / Path(sanitise_path_string(notebook_title))

        n = 1
        while notebook_path.is_dir():
            notebook_path = work_path / Path('{}_{}'.format(sanitise_path_string(notebook_title), n))
            n += 1

        notebook_media_path = Path(notebook_path / media_dir_name)
        notebook_media_path.mkdir(parents=True)

        notebook_id_to_path_index[notebook_id] = Notebook(notebook_path, notebook_media_path)

    note_id_to_title_index = {}
    converted_note_ids = []

    for note_id in config_data['note']:
        note_data = json.loads(nsx_file.read(note_id).decode('utf-8'))

        note_title = note_data.get('title', 'Untitled')
        note_ctime = note_data.get('ctime', '')
        note_mtime = note_data.get('mtime', '')

        note_id_to_title_index[note_id] = note_title

        try:
            parent_notebook_id = note_data['parent_id']
            parent_notebook = notebook_id_to_path_index[parent_notebook_id]
        except KeyError:
            continue

        print('Converting note "{}"'.format(note_title))

        content = re.sub('<img class="[^"]*syno-notestation-image-object" src=[^>]*ref="',
                         '<img src="', note_data.get('content', ''))

        attachments_data = note_data.get('attachment')
        attachment_list = []

        if attachments_data:
            for attachment_id in note_data.get('attachment', ''):

                ref = note_data['attachment'][attachment_id].get('ref', '')
                md5 = note_data['attachment'][attachment_id]['md5']
                source = note_data['attachment'][attachment_id].get('source', '')
                name = sanitise_path_string(note_data['attachment'][attachment_id]['name'])

                n = 1
                while Path(parent_notebook.media_path / name).is_file():
                    name_parts = name.rpartition('.')
                    name = ''.join((name_parts[0], '_{}'.format(n), name_parts[1], name_parts[2]))
                    n += 1

                link_path_str = '{}/{}'.format(media_dir_name, name)
                html_link_template = '<a href="{}">{}</a>'

                try:
                    Path(parent_notebook.media_path / name).write_bytes(nsx_file.read('file_' + md5))
                    attachment_list.append(html_link_template.format(link_path_str, name))
                except Exception:
                    if source:
                        attachment_list.append(html_link_template.format(source, name))
                    else:
                        print('Can\'t find attachment "{}" of note "{}"'.format(name, note_title))
                        attachment_list.append(html_link_template.format(source, 'NOT FOUND'))

                if ref and source:
                    content = content.replace(ref, source)
                elif ref:
                    content = content.replace(ref, link_path_str)

        if attachment_list:
            content = 'Attachments: {}  \n{}'.format(', '.join(attachment_list), content)
        if creation_date_in_filename and note_ctime:
            note_title = time.strftime('%Y-%m-%d ', time.localtime(note_ctime)) + note_title

        md_file_name = sanitise_path_string(note_title) or 'Untitled'
        md_file_path = Path(parent_notebook.path / '{}.{}'.format(md_file_name, 'htm'))

        n = 1
        while md_file_path.is_file():
            md_file_path = Path(parent_notebook.path / ('{}_{}.{}'.format(
                                            sanitise_path_string(note_title), n, 'htm')))
            n += 1

        md_file_path.write_text(content, 'utf-8')

        converted_note_ids.append(note_id)

    for notebook in notebook_id_to_path_index.values():
        try:
            notebook.media_path.rmdir()
        except OSError:
            pass

    not_converted_note_ids = set(note_id_to_title_index.keys()) - set(converted_note_ids)

    if not_converted_note_ids:
        print('Failed to convert notes:',
              '\n'.join(('    {} (ID: {})'.format(note_id_to_title_index[note_id], note_id)
                         for note_id in not_converted_note_ids)),
              sep='\n')

    if len(config_data['notebook']) == 1:
        notebook_log_str = 'notebook'
    else:
        notebook_log_str = 'notebooks'

    print('Converted {} {} and {} out of {} notes.\n'.format(len(config_data['notebook']),
                                                             notebook_log_str,
                                                             len(converted_note_ids),
                                                             len(note_id_to_title_index.keys())))
    try:
        recycle_bin_media_path.rmdir()
        recycle_bin_path.rmdir()
    except OSError:
        pass


input('Press Enter to quit...')
