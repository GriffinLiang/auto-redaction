# auto-redaction

### Install
* Download [Anaconda Python3](https://www.anaconda.com/products/individual) based on your system.
* Git clone this project.
* Build package: pyinstaller -F main.py
* If you encounter maximum recursion depth problem:
  * try to add the following codes into main.spec and rebuild the package with 'pyinstaller main.spec'
    ```python
    import sys
    sys.setrecursionlimit(5000)
    ```
  * Actually, I use the [solution](https://stackoverflow.com/a/60756018/4168774) which can solve this problem elegantly.
* If you encounter no module named 'pkg_resources.py2_warn', try to 
    ```
    pip install --upgrade 'setuptools<45.0.0'
    ```
    
### Usage
* put all the xlsx into input_files directory
* output_files contains the processed xlsx and log files