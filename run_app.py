import streamlit.web.bootstrap

# This is the entry point for the executable.
# It tells Streamlit to run your main app file.
if __name__ == '__main__':
    # The 'app.py' file will be bundled by PyInstaller into the executable.
    streamlit.web.bootstrap.run('app.py', '', [], flag_options={})