import pathlib
import sys
import win32api
import win32com.client
import win32con
import win32gui

def get_program_name(file_path):
	try:
		# To find where the ProductName is stored in the VersionInfo data, we first need to find the language and codepage values.
		# This code just blindly grabs the first values present. Not sure if there's a better way of doing it.
		lang, codepage = win32api.GetFileVersionInfo(file_path, "\\VarFileInfo\\Translation")[0]
	except:
		# Fallback in case VarFileInfo section doesn't exist.
		# 1033 = English US, 1252 = Windows-1252
		# I don't know where these values come from or if they apply to *all* programs that don't have VarFileInfo sections, but they apply to all the programs I tested at least.
		lang, codepage = 1033, 1252
	try:
		# Find ProductName using language and codepage values.
		program_name = win32api.GetFileVersionInfo(file_path, f"\\StringFileInfo\\{lang:04x}{codepage:04x}\\ProductName")
		
		if program_name is not None and program_name != "":
			# If ProductName was found, exit.
			return program_name	
		else:
			# If ProductName doesn't exist, try FileDescription instead.
			program_name = win32api.GetFileVersionInfo(file_path, f"\\StringFileInfo\\{lang:04x}{codepage:04x}\\FileDescription")

			if program_name is not None and program_name != "":
				# If FileDescription was found, exit.
				return program_name
	except:
		# Fallback in case StringFileInfo doesn't exist.
		pass

	# If all else fails, just use the filename.
	program_name = pathlib.Path(file_path).stem
	return program_name

def sanitise_program_name(name):
	# Remove characters that aren't valid in Windows filenames.
	invalid_chars = "<>:\"/\\|?*"
	new_name = "".join(i for i in name if i not in invalid_chars)
	return new_name

shell = win32com.client.Dispatch("WScript.Shell")
bundled = getattr(sys, 'frozen', False)
no_input = len(sys.argv) == 1

if no_input:
	# Set paths for program file.
	response = win32gui.MessageBox(0, "Create a shortcut in your 'Send to' menu?", "Send to Start Menu", win32con.MB_YESNO | win32con.MB_ICONQUESTION)
	if response == win32con.IDNO:
		sys.exit()

	file_path = str(pathlib.Path(sys.argv[0]).resolve())
	program_name = "Start menu (create shortcut)"
	dest_path = str(pathlib.Path(shell.SpecialFolders("SendTo"), program_name)) + ".lnk"
	
else:
	# Set paths for inputted file.
	file_path = str(pathlib.Path(sys.argv[1]).resolve())

	# Make sure input file exists.
	if not pathlib.Path(file_path).is_file():
		win32gui.MessageBox(0, "Input file doesn't exist.", "Send to Start Menu", win32con.MB_ICONERROR)
		sys.exit()

	startmenu_path = pathlib.Path(shell.SpecialFolders("StartMenu"), "Programs")
	program_name = sanitise_program_name(get_program_name(file_path))
	dest_path = str(pathlib.Path(startmenu_path, program_name)) + ".lnk"

folder_path = str(pathlib.Path(file_path).parent)

# Check if shortcut already exists.
if pathlib.Path(dest_path).is_file():
	if no_input:
		response = win32gui.MessageBox(0, "Shortcut already exists in 'Send to' menu. Overwrite?", "Send to Start Menu", win32con.MB_YESNO | win32con.MB_ICONWARNING)
	else:
		response = win32gui.MessageBox(0, f"Shortcut \"{program_name}\" already exists. Overwrite?", "Send to Start Menu", win32con.MB_YESNO | win32con.MB_ICONWARNING)

	if response == win32con.IDNO:
		sys.exit()
	pathlib.Path.unlink(dest_path)

# Create shortcut
shortcut = shell.CreateShortCut(dest_path)

if no_input and not bundled:
	shortcut.IconLocation = str(pathlib.Path(folder_path, "icon.ico"))
	shortcut.TargetPath = sys.executable
	shortcut.Arguments = f"\"{file_path}\""
else:
	shortcut.IconLocation = file_path
	shortcut.TargetPath = file_path

shortcut.WorkingDirectory = folder_path

shortcut.save()

if len(sys.argv) == 1:
	win32gui.MessageBox(0, "Successfully added to 'Send to' menu!", "Send to Start Menu", win32con.MB_ICONINFORMATION)
else:
	win32gui.MessageBox(0, f"Successfully added shortcut \"{program_name}\"!", "Send to Start Menu", win32con.MB_ICONINFORMATION)
