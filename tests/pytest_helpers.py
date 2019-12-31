from _pytest import junitxml


def build_filename(name, suffix='', prefix='', max_length=128):
    return '-'.join(
        filter(None, [prefix, name[:max_length - len(suffix) - len(prefix) - 5], suffix]))


def get_test_info(request, suffix='', prefix=''):
    try:
        names = junitxml.mangle_testnames(request.node.nodeid.split("::"))
    except AttributeError:
        # pytest>=2.9.0
        names = junitxml.mangle_test_address(request.node.nodeid)

    classname = '.'.join(names[:-1])
    test_name = names[-1]
    return {
        'test_name': test_name,
        'classname': classname,
    }
