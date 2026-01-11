[app]

# title of your application
title = manuscript2slides

# project root directory. default = The parent directory of input_file
project_dir = .

# source file entry point path. default = main.py
input_file = src\manuscript2slides\gui.py

# directory where the executable output is generated
exec_directory = deployment

# path to the project file relative to project_dir
project_file = 

# application icon
icon = C:\Users\zig\dev\manuscript2slides\.venv\Lib\site-packages\PySide6\scripts\deploy_lib\pyside_icon.ico

[python]

# python path
# this is where the key/value would for the python interpreter go if we wanted to include it.
# but we don't want that, because we want to deploy in a venv, and that is a local path to each machine.
# and we want check this into github, and not include absolute, local paths. because of that, when
# anyone builds, they will need to ensure that they are in a venv and it has been activated before
# running pyside6-deploy. in other words, in bash, they would run = 
#   .venv\scripts\activate  # activate venv
#   pyside6-deploy          # uses activated python automatically
# python packages to install
packages = Nuitka==2.7.11
python_path = c:\Users\zig\dev\manuscript2slides\.venv\Scripts\python.exe

[qt]

# paths to required qml files. comma separated
# normally all the qml files required by the project are added automatically
# design studio projects include the qml files using qt resources
qml_files = 

# excluded qml plugin binaries
excluded_qml_plugins = 

# qt modules used. comma separated
modules = Core,Gui,Widgets

# qt plugins used by the application. only relevant for desktop deployment
# for qt plugins used in android application see [android][plugins]
plugins = accessiblebridge,egldeviceintegrations,generic,iconengines,imageformats,platforminputcontexts,platforms,platforms/darwin,platformthemes,styles,xcbglintegrations

[nuitka]

# usage description for permissions requested by the app as found in the info.plist file
# of the app bundle. comma separated
# eg = extra_args = --show-modules --follow-stdlib
macos.permissions = 

# mode of using nuitka. accepts standalone or onefile. default = onefile
mode = onefile

# specify any extra nuitka arguments
extra_args = --quiet --noinclude-qt-translations

