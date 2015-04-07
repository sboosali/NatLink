"""
(c) 2015 Spiros Boosalis
changed a few lines to prioritize the current directory in find_natlink

Enable the NatLink Python extension to Dragon NaturallySpeaking.
$Id: enable.py 172 2008-09-29 16:33:59Z drocco $

    (c) 2008 Daniel J. Roccon
    Licensed under the Creative Commons Attribution-Noncommercial-Share Alike 3.0 United States License
    http://creativecommons.org/licenses/by-nc-sa/3.0/us/
    
Works with:
* Python installations in the 2.0 series up to 2.5 (& 1.5?).  Requires win32all
    * currently assumes Python is installed for all users...
* Dragon NaturallySpeaking versions 8 & 9; should also work with versions 6 & 7

################################################################################


Usage
=====

To install NatLink, run the script by double-clicking the file "enable.py" in
the NatLink directory.  Or from the commandline:

C:\Program Files\NatLink>enable.py

The script will attempt to automatically detect the installation locations of
Python, NaturallySpeaking, and NatLink.  If multiple versions of Python are
installed, you will be prompted to choose an installation.


To uninstall NatLink, run the following command from the commandline:

C:\Program Files\NatLink>enable.py /u


How the NatLink Installation Process Works
==========================================

The process of installing NatLink involves 1) telling Python about NatLink by updating Python Path
and 2) telling Dragon about NatLink by registering it as an application support module.  Since
NatLink is compiled as a DLL, we use the Windows DLL registration mechanics to insert the appropriate
keys into the Windows registry.  

The detailed installation steps appear below.  These steps are carried out by the installer; the documentation is provided as a reference.


0.  Install files by unzipping the NatLink distribution, for example to c:\program files\NatLink.

1.  Register NatLink dll

	cd c:\program files\NatLink\MacroSystem 
	regsvr32 natlink.dll

Each version of Python requires its own version of the NatLink DLL.  This
installer assumes that the directory .\NatLink\MacroSystem contains
precompiled versions of the DLL named for the appropriate Python version,
e.g. natlink23.dll for Python 2.3.  The installer will copy the appropriate
DLL for the version of Python to the file 'natlink.dll' before performing
the Windows registration.

As of Python 2.5, '.dll' is no longer a supported the filename extension
for modules; instead Python will search for modules with the '.pyd'
extension.  There is no binary difference between the files.  For Python
versions >= 2.5, this installer will use the new extension.

2.  Update DNS Config Files

\documents and settings\all users\app data\scansoft\ns\nssystem.ini

	[Global Clients]
add ->	.NatLink=Python Macro System

\documents and settings\all users\app data\scansoft\ns\nsapps.ini

add ->	[.NatLink]
add ->	App Support GUID={dd990001-bb89-11d2-b031-0060088dc929}

3.  Show Python where Natlink is by updating PythonPath: add path
information to the file

    ...\PythonXX\Lib\site-packages\NatLink.pth

(...\PythonXX being the Python installation directory).  The file
contains path entries to be added to Python's module search path,
sys.path, one entry per line.  For NatLink, the file will contain
the NatLink search paths:

    C:\Program Files\NatLink\MacroSystem
    C:\Program Files\NatLink\MiscScripts

The original EnableNL program and the first version of this script
used the Windows registry to set the path; thanks to Speech Computing/
Voice Code list user Reckoner for leading me to this simpler solution.


Change Log
==========

09.29.2008 	
	*   fixed a bug in the interactive Python selection process; thanks
        to Jonathan Epstein for uncovering it
    *   use .pth files in the site-packages directory to set the NatLink
	    path instead of the Windows registry


A Note on Testing
=================

This module uses the doctest framework for testing.  You can run the tests
by passing the switch '/t' to the script.  Be warned, however, that the tests
modify installation values in the registry and NaturallySpeaking config files,
so don't run the tests if this makes you uncomfortable.  Also, the tests assume
that NatLink is installed in c:\Program Files\NatLink; other installation paths
will probably cause some of the tests to fail.
"""

################################################################################
#
# imports

import os, os.path
import re
import sys
import _winreg as registry

from shutil import copyfile
from win32com.shell import shellcon, shell



################################################################################
#
# constants and globals

REGSVR_SUCCESS = 0

# NatLink site path configuration file; see the documentation for the site module
# for details <http://docs.python.org/lib/module-site.html>
NATLINK_SITE_PATH = r"Lib\site-packages\NatLink.pth"

natlink_path = None
naturallyspeaking_path = None


################################################################################
#
# operating environment

def find_natlink(search_paths=None):
    """
        Look for a NatLink installation in search_paths (and the directory
        from which this script was run) by searching for the MacroSystem
        subdirectory.

        Returns the first such installation found, or None if NatLink could
        not be found.

        >>> find_natlink()
        'C:\\\\Program Files\\\\NatLink'
    """
    global natlink_path

    if natlink_path:
        return natlink_path

    if not search_paths:
        search_paths = [os.getcwd(), r'C:\Program Files\NatLink', r'c:\NatLink']

    for path in search_paths:
        macro_path = os.path.join(path, "MacroSystem")
        if os.path.exists(macro_path) and os.path.isdir(macro_path):
            natlink_path = path
            return path

    return None        

def natlink_macro_system_path():
    return os.path.join(find_natlink(), "MacroSystem")

def natlink_misc_scripts_path():
    return os.path.join(find_natlink(), "MiscScripts")

def find_naturallyspeaking(search_paths=[r"Nuance\NaturallySpeaking10",
                                         r"Nuance\NaturallySpeaking9",
                                         r"ScanSoft\NaturallySpeaking8",
                                         r"ScanSoft\NaturallySpeaking"]):
    """
        Look for a NatLink installation in search_paths by searching for
        the nssystem.ini file.

        Returns the first such installation found, or None if NaturallySpeaking could
        not be found.

        >>> find_naturallyspeaking()
        u'C:\\\\Documents and Settings\\\\All Users\\\\Application Data\\\\ScanSoft\\\\NaturallySpeaking8'
    """
    global naturallyspeaking_path

    if naturallyspeaking_path:
        return naturallyspeaking_path
    
    application_data = shell.SHGetFolderPath(0, shellcon.CSIDL_COMMON_APPDATA, 0, 0)
    search_paths = [os.path.join(application_data, path) for path in search_paths]

    for path in search_paths:
        config_file = os.path.join(path, "nssystem.ini")
        if os.path.exists(config_file):
            naturallyspeaking_path = path
            return path

    return None

def naturallyspeaking_config_system():
    return os.path.join(find_naturallyspeaking(), "nssystem.ini")

def naturallyspeaking_config_applications():
    return os.path.join(find_naturallyspeaking(), "nsapps.ini")
        

    
################################################################################
#
# retrieve Python installation information from the registry

def get_subkey_names(reg_key):
    index = 0
    L = []
    while True:
        try:
            name = registry.EnumKey(reg_key, index)
        except EnvironmentError:
            break
        index += 1
        L.append(name)
    return L

def enumerate_python_versions():
    """
        Return a dictionary with info about installed versions of Python.

        http://mail.python.org/pipermail/python-list/2006-January/360962.html

        The dictionary keys are strings with the major and minor version number,
        e.g. '2.5'
        
        Each version value in the dictionary is represented as a tuple with 3 items:

        0   A long integer giving when the key for this version was last
              modified as 100's of nanoseconds since Jan 1, 1600.
        1   A string with major and minor version number e.g '2.4'.
        2   A string of the absolute path to the installation directory.

        Check to see that the current version of Python shows up in the registry:
        
        >>> sys.version[0:3] in enumerate_python_versions().keys()
        1
    """
    python_path = r'software\python\pythoncore'
    L = {}
    for reg_hive in [registry.HKEY_LOCAL_MACHINE]:
                      #registry.HKEY_CURRENT_USER):
        try:
            python_key = registry.OpenKey(reg_hive, python_path)
        except EnvironmentError:
            continue
        for version_name in get_subkey_names(python_key):
            key = registry.OpenKey(python_key, version_name)
            modification_date = registry.QueryInfoKey(key)[2]
            install_path = registry.QueryValue(key, 'InstallPath')
            L[version_name] = ((modification_date, version_name, install_path))
    return L


# ====================================================================
#
# file-based Python path update functions; see the documentation for the
# site module for details: http://docs.python.org/lib/module-site.html

def insert_natlink_python_path(python_version=None, python_install_path=None):
    """
        Create a .pth file in the site-packages directory containing NatLink
        path information.
    """
    if python_version:
        python = enumerate_python_versions()[python_version]
        python_install_path = python[2]

    filename = os.path.join(python_install_path, NATLINK_SITE_PATH)
    paths = "\n".join( (natlink_macro_system_path(), natlink_misc_scripts_path()) )

    write_string_to_file(filename, paths)


def delete_natlink_python_path():
    """
        Delete NatLink.pth files from the Python site-packages directory,
        thereby removing NatLink from the Python search path.
    """
    for python, details in enumerate_python_versions().items():
        try:
            filename = os.path.join(details[2], NATLINK_SITE_PATH)
            os.remove(filename)
        except:
            pass


# ====================================================================
#
# registry-based Python path update functions: to add NatLink to the Python path
# via the registry, add key
#
#    HKEY_LOCAL_MACHINE\Software\Python\PythonCore\XXX\PythonPath\NatLink
#
# where XXX is the Python version string, e.g. '2.5', and set the default value to
# the NatLink scripts paths:
#
#    "***\MacroSystem;***\MiscScripts"
#
# where *** is the current path to the NatLink subsystem.  
#
# e.g. "C:\Program Files\NatLink\MacroSystem;C:\Program Files\NatLink\MiscScripts"
#

def get_natlink_python_path_registry(python_version):
    """
        Given a Python version string, e.g. '2.5', returns the value of the
        registry key:

            HKEY_LOCAL_MACHINE\Software\Python\PythonCore\XXX\PythonPath\NatLink

        or None if it does not exist.

        >>> insert_natlink_python_path_registry('2.5')
        >>> get_natlink_python_path_registry('2.5')
        u'C:\\\\Program Files\\\\NatLink\\\\MacroSystem;C:\\\\Program Files\\\\NatLink\\\\MiscScripts'
    """
    try:
        key = registry.OpenKey(registry.HKEY_LOCAL_MACHINE,
                               r'Software\Python\PythonCore\%s\PythonPath\NatLink'
                               % python_version)
        return registry.EnumValue(key, 0)[1]
    except:
        return None

def delete_natlink_python_path_registry(python_version):
    """
        Given a Python version string, e.g. '2.5', deletes the registry key:

            HKEY_LOCAL_MACHINE\Software\Python\PythonCore\XXX\PythonPath\NatLink

        Returns true if the delete was successful, false otherwise.            

        >>> get_natlink_python_path_registry('2.5')
        u'C:\\\\Program Files\\\\NatLink\\\\MacroSystem;C:\\\\Program Files\\\\NatLink\\\\MiscScripts'
        >>> delete_natlink_python_path_registry('2.5')
        True
        >>> get_natlink_python_path_registry('2.5') is None
        True
    """
    try:
        registry.DeleteKey(registry.HKEY_LOCAL_MACHINE,
                         r'Software\Python\PythonCore\%s\PythonPath\NatLink'
                         % python_version)
        return True
    except:
        return False

def insert_natlink_python_path_registry(python_version):
    """
        Given a Python version string, e.g. '2.5', sets the value of the registry key

            HKEY_LOCAL_MACHINE\Software\Python\PythonCore\XXX\PythonPath\NatLink

        to

            "%s;%s" % (natlink_macro_system_path(), natlink_misc_scripts_path())

        creating the key if it does not already exist.            

        >>> delete_natlink_python_path_registry('2.5')
        True
        >>> get_natlink_python_path_registry('2.5') is None
        True
        >>> insert_natlink_python_path_registry('2.5')
        >>> get_natlink_python_path_registry('2.5')
        u'C:\\\\Program Files\\\\NatLink\\\\MacroSystem;C:\\\\Program Files\\\\NatLink\\\\MiscScripts'
    """
    if get_natlink_python_path(python_version) is None:
        registry.SetValue(registry.HKEY_LOCAL_MACHINE,
                          r'Software\Python\PythonCore\%s\PythonPath\NatLink'
                              % python_version,
                          registry.REG_SZ,
                          "%s;%s" % (natlink_macro_system_path(), natlink_misc_scripts_path()))
        

    
################################################################################
#
# register/unregister the NatLink DLL

def unregister_natlink_module(base=natlink_macro_system_path()):
    modules = ['natlink.dll', 'natlink.pyd']
    modules = [os.path.join(base,module) for module in modules]
    modules = [module for module in modules if os.path.exists(module)]
    result = True

    for module in modules:
        result = result & (os.system('regsvr32 /u /s "%s"' % (module)) == REGSVR_SUCCESS)

    return result        

def register_natlink_module(python_version,base=natlink_macro_system_path()):
    python = python_version.replace ('.', '')
    natlink = os.path.join(base, 'natlink%s')
    extension = '.dll'

    if int(python) >= 25:
        extension = '.pyd'
        
    source_module = natlink % (python + extension)
    destination_module= natlink % (extension)

    copyfile(source_module,destination_module)
    return (os.system('regsvr32 /s "%s"' % (destination_module)) == REGSVR_SUCCESS)
    
        
            
            
    
################################################################################
#
# add/remove NaturallySpeaking ini file entries

def file_as_string(filename):
    f = open(filename, 'r')
    result = f.read()
    f.close()
    return result

def write_string_to_file(filename, contents):
    f = open(filename, 'w')
    f.write(contents)
    f.close()
    
def remove_naturallyspeaking_config_entries():
    """
        Remove the NatLink configuration settings from the NaturallySpeaking
        config files.  NaturallySpeaking does not like the output format that
        ConfigParser uses, so while it is still useful for testing, modifications
        to the NaturallySpeaking ini files must be done manually.

        >>> import ConfigParser
        >>> remove_naturallyspeaking_config_entries()
        >>> system_config = ConfigParser.RawConfigParser()
        >>> foo = system_config.read(naturallyspeaking_config_system())
        >>> system_config.has_option("Global Clients", ".NatLink")
        False
        >>> apps_config = ConfigParser.RawConfigParser()
        >>> foo = apps_config.read(naturallyspeaking_config_applications())
        >>> apps_config.has_section(".NatLink")
        False
    """
    system_config = file_as_string(naturallyspeaking_config_system())
    system_config = re.sub(r"\s*.NatLink[^\n]*\n", "\n", system_config)

    apps_config = file_as_string(naturallyspeaking_config_applications())
    apps_config = re.sub(r"\s*App Support GUID\s*=\s*\{dd990001-bb89-11d2-b031-0060088dc929\}\s*\n",
                         "\n",
                         apps_config)
    apps_config = re.sub(r"\s*\[.NatLink\]\s*\n", "\n", apps_config)
    
    write_string_to_file (naturallyspeaking_config_system(), system_config)
    write_string_to_file (naturallyspeaking_config_applications(), apps_config)

def add_naturallyspeaking_config_entries():
    """
        Add the NatLink configuration settings to the NaturallySpeaking
        config files.  NaturallySpeaking does not like the output format that
        ConfigParser uses, so while it is still useful for testing, modifications
        to the NaturallySpeaking ini files must be done manually.

        >>> import ConfigParser
        >>> add_naturallyspeaking_config_entries()
        (nssystem.ini)  (nsapps.ini)
        >>> system_config = ConfigParser.RawConfigParser()
        >>> foo = system_config.read(naturallyspeaking_config_system())
        >>> system_config.has_option("Global Clients", ".NatLink")
        True
        >>> apps_config = ConfigParser.RawConfigParser()
        >>> foo = apps_config.read(naturallyspeaking_config_applications())
        >>> apps_config.has_section(".NatLink")
        True
    """
    print "(nssystem.ini) ",
    system_config = file_as_string(naturallyspeaking_config_system())
    if not re.findall(r"\s*.NatLink[^\n]*\n", system_config):
        system_config = re.sub(r"\[Global Clients\][^\n]*\n",
                               "[Global Clients]\n.NatLink=Python Macro Subsystem\n",
                               system_config)

    print "(nsapps.ini) ",
    apps_config = file_as_string(naturallyspeaking_config_applications())
    if not re.findall(r"\s*\[.NatLink\]\s*\n", apps_config):
        apps_config += "\n[.NatLink]\nApp Support GUID={dd990001-bb89-11d2-b031-0060088dc929}\n\n"

    write_string_to_file(naturallyspeaking_config_system(), system_config)
    write_string_to_file(naturallyspeaking_config_applications(), apps_config)
        

    
################################################################################
#
# interactive console installation helpers

def find_and_select_python_installation(pythons=enumerate_python_versions().keys()):
    """
        Returns a string containing the major and minor version number of an
        installed Python installation as returned by enumerate_python_versions()

        If only one version of Python is installed, its version number is returned.

        If more than one version of Python is installed, the user chooses which
        version they want.
    """
    if len(pythons) == 1:
        print "Found Python installation (%s)." % (pythons[0])
        return pythons[0]
    else:
        # 09.29.2008 00:14 djr: FIXED: interactive selection wasn't working
        #   because the return was missing below... oops :-}
        #   thanks to Jonathan Epstein for uncovering this (embarrassing!) bug
        return select_python_installation(pythons)

def select_python_installation(pythons,_input_function=raw_input):
    """
        Given a list of Python installations, interactively select one using
        a console menu.  The function will accept a menu choice (0,1,2,...)
        or a version string ('2.5') from the user.

        pythons: a list of Python versions strings, e.g. ['2.3', '2.5']

        doctests:
        >>> select_python_installation(['2.3', '2.5'],lambda(x): '0')
        Choose an installation:...
        '2.5'
        >>> select_python_installation(['2.3', '2.5'],lambda(x): '')
        Choose an installation:...
        '2.5'
        >>> select_python_installation(['2.3', '2.5'],lambda(x): '1') 
        Choose an installation:...
        '2.3'
        >>> select_python_installation(['2.3', '2.5'],lambda(x): '2.5')
        Choose an installation:...
        '2.5'
        >>> select_python_installation(['2.3', '2.5'],lambda(x): '2.3') 
        Choose an installation:...
        '2.3'
    """
    pythons.sort()
    pythons.reverse()

    print "Choose an installation:\n"
    for index in range(len(pythons)):
        print "%d. %s" % (index, pythons[index])

    choice = _input_function("\nWhich Python version would you like? [0] ")

    if choice in pythons:
        selection = choice
    elif choice == '':
        selection = pythons[0]
    else:    
        try:
            selection = pythons[int(choice)]
        except:
            print "Couldn't understand selection: %s" % (choice)
            return select_python_installation(pythons, _input_function)
        
    return selection
        


################################################################################
#
# script entry point for installation and testing


def install_interactive(registry=False):
    """
        Install NatLink, prompting the user if necessary.  This method inserts
        the NatLink configuration into NaturallySpeaking's config files, adds
        NatLink to PythonPath, and registers the NatLink DLL.

        By default, this method connects Python to NatLink using a .pth file in
        Python's site-packages directory.  Set the registry parameter to True to
        install NatLink to the Python path in the Windows registry.
    """
    python = find_and_select_python_installation()

    print "Unregistering NatLink module"
    unregister_natlink_module()
    
    print "Registering NatLink module for Python %s" % (python)
    register_natlink_module(python)

    print "Adding NatLink to PythonPath for Python %s" % (python)
    if not registry:
        insert_natlink_python_path(python)
    else:
        insert_natlink_python_path_registry(python)

    print "Adding NatLink to Dragon NaturallySpeaking config files"
    add_naturallyspeaking_config_entries()

    print "NatLink installation complete."

def uninstall_interactive(registry=False):
    """
        Uninstall NatLink, prompting the user if necessary.  This method removes
        the NatLink configuration from NaturallySpeaking's config files, removes
        NatLink from PythonPath, and un-registers the NatLink DLL.

        By default, this method assumes that Python is connected to NatLink using
        a .pth file in site-packages.  Set the registry parameter to True to
        uninstall NatLink from the Python path in the Windows registry.
    """
    print "Removing NatLink from Dragon NaturallySpeaking config files"
    remove_naturallyspeaking_config_entries()

    print "Removing NatLink PythonPath"
    if not registry:
        delete_natlink_python_path()
    else:
        for python in enumerate_python_versions().keys():
            if delete_natlink_python_path_registry(python):
                print python + " ",
    
    print "Unregistering NatLink module"        
    unregister_natlink_module()
    
    print "NatLink removal complete."    


def install_registry():
    install_interactive(True)
    
def uninstall_registry():
    uninstall_interactive(True)

    
def _test():
    import doctest
    doctest.testmod(verbose=1, optionflags=(doctest.ELLIPSIS |
                                            doctest.REPORT_NDIFF |
                                            doctest.NORMALIZE_WHITESPACE))


if __name__ == "__main__":
    commands = {
        "/t": _test,
        "/i": install_interactive,
        "/u": uninstall_interactive,
        "/ir": install_registry,
        "/ur": uninstall_registry
    }

    if len(sys.argv) <= 1:
        command = commands["/i"]
    elif not sys.argv[1] in commands.keys():
        print "usage: %s [/i | /u | /ir | /ur]" % (sys.argv[0])
        print("\t/i | /u\t\tInstall/uninstall NatLink using the site module \
                \t\tto update the Python path")
        print("\t/ir | /ur\tInstall/uninstall NatLink using the Windows \
                \t\t\tregistry to update the Python path")
        sys.exit(-1)
    else:
        command = commands[sys.argv[1]]
        
    command()
