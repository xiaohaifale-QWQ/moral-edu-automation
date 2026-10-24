# -*- coding: utf-8 -*-
"""打印文件的 sha256 与字节数（发布 SOP 用）。

用法：python tools/hash.py <文件路径>
"""

import hashlib
import os
import sys


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    path = sys.argv[1]
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    print(f"文件：{path}")
    print(f"sha256：{h.hexdigest()}")
    print(f"字节数：{os.path.getsize(path)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
