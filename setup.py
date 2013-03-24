from distutils.core import setup

setup(name='l5x',
      version='0.1',
      author='Jason Valenzuela',
      author_email='jvalenzuela1977@gmail.com',
      packages=['l5x', 'tests'],
      url='http://pypi.python.org/pypi/l5x/',
      license='LICENSE.txt',
      description='RSLogix .L5X interface.',
      long_description=open('README.txt').read(),
      classifiers=[
        'Development Status :: Production/Stable',
        'License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)',
        'Intended Audience :: Developers',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 2.7',
        'Topic :: Text Processing :: Markup :: XML'
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Scientific/Engineering :: Interface Engine/Protocol Translator'
        ])