import ctypes
import os
import subprocess

with open('Lab_0.txt', 'wb', 0) as file:
    subprocess.run('ipconfig /all | findstr /i "IPv4"', stdout=file, shell=True, check=True)

with open('Lab_0.txt', 'r', encoding="cp866") as file:
    print(file.read())


