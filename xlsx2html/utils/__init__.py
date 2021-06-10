import ctypes


def hash_str(s: str):
    return ctypes.c_uint64(hash(s)).value.to_bytes(8, "big").hex()
