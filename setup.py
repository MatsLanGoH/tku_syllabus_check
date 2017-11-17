from distutils.core import setup

setup(
    name='TkuSyllabusCheck',
    version='0.1.2',
    author='Matthias Immo Lambrecht',
    author_email='mats.lan.pod@googlemail.com',
    packages=['syllabus_check'],
    license='LICENSE.txt',
    description='Syllabus checker for Campusmate files',
    long_description=open('README.txt').read(),
    install_requires=[
        "lxml==4.1.0",
        "python_docx == 0.8.6",
        "xlrd == 1.1.0",
    ],
)
