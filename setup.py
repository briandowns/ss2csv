from distutils.core import setup

setup(
    name='ss2csv',
    version='0.0.1',
    author='Brian Downs',
    author_email='brian.downs@gmail.com',
    maintainer='Brian Downs',
    maintainer_email='brian.downs@gmail.com',
    scripts=['ss2csv/ss2csv.py'],
    packages=['ss2csv'],
    package_data={
        '': ['*.txt'],
    },
    url='https://github.com/briandowns/ss2csv',
    license='Apache',
    description='Quickly save spreadsheet workbooks to csv files.',
    long_description=open('README.md').read(),
)
