from setuptools import setup

setup(
    name="ExDForm",
    version="1.0",
    author="exact_data",
    description="Application de creation d'un formulaire de saisie dynamique de donnée",
    options={
        'build_exe': {
            'includes': [
                'PyQt6', 'pandas', 'docx',
                'scipy', 'sqlite3', 'numpy'
            ],
            'include_files': ['assets/']
        }
    }
)