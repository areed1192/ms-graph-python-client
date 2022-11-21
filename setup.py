from setuptools import setup
from setuptools import find_namespace_packages

# Open the README file.
with open(file="README.md", mode="r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="ms-graph-python-client",
    author="Alex Reed",
    author_email="coding.sigma@gmail.com",
    version="0.1.0",
    description="A Python Client Application that allows interaction with the Microsoft Graph API.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/areed1192/ms-graph-python-client",
    install_requires=["requests", "msal"],
    packages=find_namespace_packages(include=["ms_graph", "ms_graph.*"]),
    python_requires=">3.8",
)
