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

    text_owner = str(name_dict)
    owner = text_owner.split(':')[0][2:-1]
    doc = text_owner.split(':')[1][2:-2]
    return path, owner, doc


def write_in_owner(path, owner, doc):
    with open(path + '\\' + 'owner_doc.txt', 'w') as f:
        f.write(owner + '\n' + doc)


def main():
    path, owner, doc = create_dict_f()
    write_in_owner(path, owner, doc)


if __name__ == '__main__':
    main()
