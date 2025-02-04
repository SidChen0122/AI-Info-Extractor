# -*- coding: utf-8 -*-

import os
import sys
import json
import datetime
import re
import xlsxwriter
import xlrd
from openai import OpenAI

# file must exist and not empty
def exist(p):
    if not os.path.exists(p):        
        # if targeted file is txt, create this file for user's convenience
        if p.split('.')[-1] == 'txt':            
            print(f'{p} doesn\'t exist. \n A new {p} has been created, please paste required content in.')
            f = open(p, mode='w', encoding='utf-8-sig')
            f.write('')
            f.close()
            input('\n* Any key to exist, please run the script again afterwards')
        else:
            input('"{}"文件不存在，请确认，任意键退出'.format(p))
    if p.split('.')[-1] == 'txt':
        f = open(p, mode='r')
        if len(f.read())<1:
            input(f'{p} is empty, please check')
        f.close()
    exit()

# file_format_list = [‘.xls’, ‘.xlsx’]
def file_name(file_dir, file_format_list):
    """Find goal file path"""
    l = []
    # root: current path, dirs
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if os.path.splitext(file)[1] in file_format_list:
                l.append(os.path.join(root, file))
    return l

def human_choose(choices, times):
	# choices: 供选择的list
	# times： 需要选择的个数
    id = []
    for i, j in enumerate(choices):
        print(i, ':', j)
        id.append(i)
    print(f'本次需选择{times}个数据，请输入对应编号：\n（严格按照目标字段顺序先后输入，英文逗号分隔，Enter确定）')
    file = []
    input_check(id, len(id), file, times)
    return file

def input_check(must_in, l, f, count):
    # three chances
    for i in range(4):
        if i == 3:
            input('请重新运行脚本，任意键退出')
            exit()
        r = input()
        # numbers limit
        if len(r) == 0:
            print('输入值为空，剩余输入次数为:', (2 - i))
            continue
        elif len(r) > l:
            print('输入值超出限制，剩余输入次数为:', (2 - i))
            continue
        else:
            rr = r.split(',')
            # avoid same number
            g = 'out'
            rrr = []
            for j in rr:
                # must fit the set mode
                for k in must_in:
                    if int(j) == k:
                        g = 'in'
                        print(j)
                        rrr.append(j)
                        break
                    else:
                        g = 'out'
            if len(rrr) != count:
                g = 'out'
            if g == 'out':
                print('输入错误，剩余输入次数为:', (2 - i))
                continue
            else:
                for m in rrr:
                    print('input:', m)
                    f.append(int(m))
                break


# open a xlsx file and read data
def excel_read(p, select, f):
    # pip install xlrd==1.2.0
    exist(p)
    print(f'打开：{p}')
    data = xlrd.open_workbook(p)
    table = data.sheets()[0]
    lines = table.nrows
    # select specific columns
    if len(select) != 0:
        head = table.row_values(0)
        select_c = []
        for i, j in enumerate(head):
            if j in select:
                select_c.append(i)
                if len(select_c) == len(select):
                    break
        if len(select_c) != len(select):
            print(f'文件{p}中未找到目标字段{select}，请确认！\n')
            select_c = human_choose(head, len(select))
            print('选择列：', select_c)
        f.append(select)
        for i in range(lines-1):
            row_s = []
            for j in select_c:
                row_s.append(table.cell_value(i+1, j))
            f.append(row_s)
    else:
        for i in range(lines):
            f.append(table.row_values(i + 1))
    print(f'共读取到：{len(f)} 行数据（含首行）')
    return


# write data into txt
def txt_write(path, data):
    f = open(path, mode='a', encoding='utf-8')
    data_line = []

    def list_split(d):
        if type(d) in (list, tuple): 
            for m in d:
                list_split(m)
        else:
            data_line.append(d)
        return

    if type(data) in (list, tuple):
        for i in data:
            list_split(i)
            j = ''
            for k in data_line:
                j += ','+str(k)
            f.write(j.lstrip('"",')+'\n')
            data_line.clear()
    else:
        f.write(data)
    print('保存至：'+path)
    f.close()

    # file must existence
def exist(p):
    if not os.path.exists(p):
        input('"{}"文件不存在，请确认，任意键退出'.format(p))
        exit()

def txt_read(path):
    """read txt file and return a list"""
    exist(path)
    f = open(path, mode='r', encoding='utf-8')
    l = []
    wrong = 0
    for i in f.readlines():
        j = i.strip('\n')
        if len(j) != 0:
            l.append(j)
        else:
            wrong += 1
    print('* Have read the file: {}; empty lines : {}'.format(path, wrong))
    return l

def excel_write(ph, data):
    from math import ceil
    # create a table and write in data
    table = xlsxwriter.Workbook(ph)
    # avoid the rows limit for xlsx:1048576\xls:65536
    table_type = ph.split('.')[-1]
    if table_type == 'xlsx':
        # ceil: 向上取整
        sheet_num = ceil((len(data)-1)/1048000)
    else:
        sheet_num = int(len(data)/6500)
    print(f'* Total Data:{len(data)} rows; Target Sheet:{sheet_num} sheets\n* writing......')
    for s in range(sheet_num):
        sheet = table.add_worksheet(f'sheet{s}')
        if table_type == 'xlsx':
            id_begin = s * 1048000
            id_end = 1048000 + id_begin
        else:
            id_begin = s * 6500
            id_end = 6500 + id_begin
        if id_end > len(data):
            id_end = len(data)-1
        # write head
        sheet.write(0, 0, 'ID')
        for k in range(len(data[0])):
            sheet.write(0, k + 1, data[0][k])
        # write data
        for i in range(id_begin, id_end):
            # write ID column
            sheet.write(i-id_begin+1, 0, i + 1)
            for j in range(len(data[i + 1])):
                # 这是最宽泛的将数字、字符串分开的方法，不过有可能出错，可以考虑结合“判断是否为数字”的代码使用
                try:
                    sheet.write(i-id_begin+1, j + 1, data[i + 1][j])
                except:
                    sheet.write(i-id_begin+1, j + 1, str(data[i + 1][j]))
        print(f'* Have finished:sheet{s}')
    table.close()
    print('###save as ' + ph)

def find_txt_files (directory, file_pre, file_type = '.txt'):
    pre_len = len(file_pre)
    txt_files = []
    for file in os.listdir(directory):
        if file.endswith(file_type) and os.path.basename(file)[0:pre_len] == file_pre:
            txt_files.append(file)
    return txt_files

def dict_extract(content, prefix=""):
        """
        Recursively extracts information from a nested dictionary.
        
        Args:
            content (dict): The dictionary to process.
            prefix (str): Prefix for lowest level of keys to save as headings.

        Returns:
            list: A list of tuples containing headings and corresponding cell values.
        """
        data = []
        for key, value in content.items():
            current_prefix = f"{prefix}_{key}" if prefix else key
            if isinstance(value, dict):
                data.extend(dict_extract(value, current_prefix))
            else:
                data.append((current_prefix.split('_')[-1], value))
        return data # [('Post ID', '140'), ('Relevance', 'No'), ...]

def batch_prepare():
    print('\n* Please select data source')
    mode = human_choose(['load prompts from txt (general prompts, one prompt per line)', 
                        'load prompts from Xiaohongshu data (acquired by RedNoteSpider script)'], 1)[0]
    if mode == 0:
        print('\n* load prompts from prompts.txt (general prompts, one prompt per line)')
        exist('prompts.txt')
        post_list = txt_read('prompts.txt')
    else:
        print('\n* load prompts from Xiaohongshu data')
        print('\n* list all excel files in this folder')
        p1 = os.getcwd()
        p = file_name(p1, ['.xls', '.xlsx'])
        file_path = p[human_choose(p, 1)[0]]
        print('\n* read the sheet0 data')
        info_targeted = ['ID', 'ip', 'post_date', 'author', 'title', 'content', 'comments_selected']
        info_list = []
        excel_read(file_path, info_targeted, info_list)
        print('\n* convert to prompts for ChatGPT Model:\nID: {ID} title: {title} posted by {author} on {post_date}  in {ip} with content: {content} with comments: {comments_selected}')
        post_list = []
        for i in range(len(info_list)-1):
            prompt = f'ID: {int(info_list[i+1][0])} title: {info_list[i+1][4]} posted by {info_list[i+1][3]} on {info_list[i+1][2]} in {info_list[i+1][1]} with content: {info_list[i+1][-2]} with comments: {info_list[i+1][-1]}'
            post_list.append(prompt)
        print(f'Save {len(post_list)} prompt(s) to submit')
        dt = datetime.datetime.now().strftime('%d%m%Y-%H%M%S')
        txt_write(f'prompts_{dt}.txt', post_list)

def batch_submit():    
    # combine posts and extraction requirements, and prepare the jsonl file
    prompt_in = 'prompt_your_requirement.txt'
    input(f'''\n* Read your prompt (command for ChatGPT to execute)
        \nPlease save all your requirements in {prompt_in}
        \nMUST INCLUDE A 'Post ID' AS THE FIRST ITEM''')
    exist(prompt_in)
    # read user requirement
    f = open(prompt_in, mode='r', encoding='utf-8-sig')
    prompt_up = f.read()
    f.close()
    print('\n* Received your requirements:\n\n', prompt_up)
    
    # read posts info
    print('\n* convert all your prompts and your requirements to a json file for the batch')
    requests = []
    prompts_file = find_txt_files(base_dir, file_pre = 'prompts')
    if len(prompts_file) == 1:
        prompts_targeted = prompts_file[0]
    elif len(prompts_file) > 1:
        print('\n* found multi prompts files, please choose one')
        prompts_targeted = prompts_file[human_choose(prompts_file, 1)[0]]
    else:
        input('\n* Cannot find any prompts file to process, please check (any key to exit)')
        exit()
    post_list = txt_read(prompts_targeted)
    
    # process the posts
    jsonl = f'batchinput_{dt}.jsonl'
    for i in range(len(post_list)):
        request = {"custom_id": f"request-{i}", "method": "POST", "url": "/v1/chat/completions", 
                "body": {"model": "gpt-4o-mini", 
                            "messages": [{"role": "assistant", "content": prompt_up},{"role": "user", "content": str(post_list[i])}]}}
        requests.append(json.dumps(request))
    txt_write(jsonl, requests)

    # upload the batch
    print('\n* Uploading the file to ChatGPT...')
    try:
        batch_input_file = client.files.create(
            file=open(jsonl, "rb"),
            purpose="batch"
            )
        print('\n* file has been uploaded')
    except:
        input('\n* Please check your network on ChaptGPT connection (any key to exit)')
        exit()    
    
    # start the batch
    print('\n* Start the batch')
    batch_task = client.batches.create(
        input_file_id=batch_input_file.id,
        endpoint="/v1/chat/completions",
        completion_window="24h",
        metadata={
            "description": "nightly eval job"
        }
    )
    print('\n* All Done! The processing usually need a whole day (late night batch for cheaper price)')
    batch_log = [('bach file (.jsonl)', 'batch uploaded id', 'batch task id')]
    batch_log.append((jsonl, batch_input_file.id, batch_task.id))
    txt_write(f'batch_log_{dt}.txt', batch_log)

def batch_status():
    print('\n* Checking the batch status...')
    batch_log_lists = find_txt_files(base_dir, file_pre = 'batch_log')
    if len(batch_log_lists) == 1:
        batch_targeted = batch_log_lists[0]
    elif len(batch_log_lists) > 1:
        print('\n* found multi batch tasks, please choose one')
        batch_targeted = batch_log_lists[human_choose(batch_log_lists, 1)[0]]
    else:
        input('\n* Cannot find any batch file to process, please check (any key to exit)')
        exit()
    batch_id = txt_read(batch_targeted)[-1].split(',')[-1]

    try:
        batch = client.batches.retrieve(batch_id)
        print('\nSTATUS: ', batch.status)
        if batch.status == 'completed':
            batch_retrieve(batch.output_file_id)
    except:
        input('\n* Please check your network on ChaptGPT connection (any key to exit)')
        exit()

def batch_retrieve(output_file_id):
    print('\n* Retrieving batch results...')
    file_response = client.files.content(output_file_id)
    txt_write('results_raw.txt', file_response.text)

    try:
        print('\n* Format and save batch results as excel file...')
        # save responses as list
        responses = file_response.text.rstrip('\n').split('\n')
        print(len(responses), 'response(s)\n' )

        # Initialize variables
        results_dict = {'created_time':{}, 'model_used':{}}
        results_list = []
        headings_final = ['created_time', 'model_used']

        # extract key information
        for i in responses:
            i = json.loads(i)
        
            timestamp = int(i['response']['body']['created'])
            created_time = datetime.datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')

            model_used = i['response']['body']['model']    

            # extract content
            content_raw = i['response']['body']['choices'][0]['message']['content']
            match = re.search(r'{.*}', content_raw, re.DOTALL)
            content = json.loads(match.group(0))
            headings, cells = zip(*dict_extract(content)) # ['Post ID', 'Relevance', ...], ['140', 'No', ...]
            
            # format content
            post_id = str(cells[0])
            results_dict['created_time'][post_id] = created_time
            results_dict['model_used'][post_id] = model_used
            for i in range(len(headings)):
                heading = headings[i]
                if heading not in results_dict.keys():          
                    results_dict[str(heading)] = {}    
                results_dict[str(heading)][post_id] = cells[i] # {'created_time':{}, 'model_used':{}, 'Post ID': {'140': '140', ...}, 'Relevance': {'140': 'No'...}, ...}

        # convert result_dict to a list for exporting
        headings_final = list(results_dict.keys()) # ['created_time', 'model_used', 'Post ID', 'Relevance', ...]
        results_list.append(headings_final)
        for post_id in results_dict['Post ID'].keys():
            # save info of each post
            post = []
            for i in range(len(headings_final)):         
                info_key = headings_final[i]
                info_list = results_dict[info_key]
                if str(post_id) in info_list:
                    info = results_dict[info_key][str(post_id)]
                else:
                    info = ''  
                post.append(info)
            results_list.append(post)

        # save as excel file    
        excel_write(f'results_{dt}.xlsx', results_list)
    except:
        print('\n* unexpected content format, maybe you need to process the returns saved above manually')

def sub_script():
    print('\n* Please choose the script you want to execute')
    scripts = ['Read posts and save as json file for batch API', '(re)Submit the batch',
                'Check batch status and Retrieve results', 'Exit?']
    mode = human_choose(scripts, 1)[0]
    if mode == 0:
        batch_prepare()
        sub_script()
    elif mode == 1:
        batch_submit()
        sub_script()
    elif mode == 2:
        batch_status()
        sub_script()
    else:
        exit()

if __name__ == '__main__':
    # script info
    print('* 本脚本用于调用ChatGPT-4o-mini模型结构化提取文本关键信息')
    # check and download the updated progress
    progress = {'gpt_info_extraction.exe': ['20241206', 'https://chdeducn-my.sharepoint.com/:f:/g/personal/201541020106_chd_edu_cn/ErR_a9ObbMlNrwM5yHFjQIUBj0DWDU8S3LHYsoN50ZXDeA?e=Carwlr'
                                    ]
                , 'UpdateLogs.txt': ['20241206', 'https://chdeducn-my.sharepoint.com/:f:/g/personal/201541020106_chd_edu_cn/ErR_a9ObbMlNrwM5yHFjQIUBj0DWDU8S3LHYsoN50ZXDeA?e=Carwlr'
                                     ]
                }
    # url must be the direct sharing link if it's from OneDrive where you can check latest update date
    url_update = 'https://chdeducn-my.sharepoint.com/:f:/g/personal/201541020106_chd_edu_cn/ErR_a9ObbMlNrwM5yHFjQIUBj0DWDU8S3LHYsoN50ZXDeA?e=Carwlr'
    for i in progress:
        print(f'* {i}，版本号{progress[i][0]}，')
    print('*** 有限技术支持：sidchen0 @ qq.com ***')

    # initial date for file names
    dt = datetime.datetime.now().strftime('%d%m%Y-%H%M%S')
    
    # API check
    gpt_api = '''\n* Please define ChatGPT API by
      \nrunning this code in any terminal: setx OPENAI_API_KEY "your_api_key_here"
      \nand then restart the script
      \n guidance: 
      a) win+R; 
      b) type in \'cmd\'; 
      c) paste the code provided above in the pop-up terminal , it is expected to see the result: \'SUCCESS: Specified value was saved.\''''
    try:        
        client = OpenAI()
    except:
        input(gpt_api)
        exit()
    
    # Change working directory to the same folder as the script
    # Determine the base directory of the executable or script
    if getattr(sys, 'frozen', False):  # If running as a bundled executable
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    # Change the working directory to the base directory
    os.chdir(base_dir)

    # initial script
    sub_script()