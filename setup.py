from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="testcase-converter",
    version="0.1.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="Convert test cases between Excel and XMind formats",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/misldy/TestCaseConverter",
    packages=find_packages(),
    include_package_data=True,
    package_data={
        'testcase_converter': ['resources/*.xmind']
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.7',
    install_requires=[
        'xmind',
        'openpyxl',
        'dataclasses; python_version<"3.7"'
    ],
    entry_points={
        'console_scripts': [
            'testcase-converter=testcase_converter.converter:main'
        ]
    }
)