from setuptools import setup, find_packages

setup(
    name="gitcli",
    version="1.0",
    py_modules=["gitcli"],
    install_requires=[],
    entry_points={
        "console_scripts": [
            "pushi=gitcli:main",
        ],
    },
)
