import os


import win32security
from os.path import isfile, join


def create_dict_f():
    path = os.path.abspath('')
    list_text = [f for f in os.listdir(path) if isfile(join(path, f))]
    name_dict = {}
    for i in range(len(list_text)):
        if list_text[i].split('.')[1] in ('xlsx', 'xls', 'doc', 'ppt', 'pptx', 'docx'):
            path_doc = path + '\\' + list_text[i]
            sd = win32security.GetFileSecurity(path_doc, win32security.OWNER_SECURITY_INFORMATION)
            owner_sid = sd.GetSecurityDescriptorOwner()
            name, domain, type = win32security.LookupAccountSid(None, owner_sid)
            try:
                name_dict[name] += [list_text[i]]
            except KeyError:
                name_dict[name] = [list_text[i]]

    return path, name_dict


def write_in_owner(path, text_owner):
    with open(path + '\\' + 'owner_doc.txt', 'w') as f:
        for k in text_owner:
            string_doc = text_owner[k]
            f.write(f"{k} \n")
            for i in range(len(string_doc)):
                f.write(f'{string_doc[i]} \n')
            f.write('\n')


def main():
    path, name_dict = create_dict_f()
    write_in_owner(path, name_dict)


if __name__ == '__main__':
    main()
