

def test_add_group(app):
    old_list = app.groups.get_group_list()
    item = app.groups.read_excel()
    app.groups.add_new_group(item)
    new_list = app.groups.get_group_list()
    old_list.append(item)
    assert  sorted(old_list) == sorted(new_list)


def test_delete_group(app):
    old_list = app.groups.get_group_list()
    app.groups.delete_new_group(old_list[1])
    new_list = app.groups.get_group_list()
    old_list.remove(old_list[1])
    assert  sorted(old_list) == sorted(new_list)
