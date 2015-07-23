from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
buildOptions = dict(packages = [], excludes = [])

base = 'Console'

executables = [
    Executable('script.py', base=base)
]

setup(name='cerema',
      version = '1.0',
      description = 'Cr\xc3\xa9ation de rapport de point de mesure',
      options = dict(build_exe = buildOptions),
      executables = executables)
