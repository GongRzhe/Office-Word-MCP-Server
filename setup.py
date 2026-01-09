from setuptools import setup, find_packages

setup(
    name="word-document-server",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "fastmcp>=2.8.1",
        "python-docx>=0.8.11",
        "lxml>=4.9.0",
        "python-dotenv>=1.0.0",
    ],
    python_requires=">=3.8",
)
