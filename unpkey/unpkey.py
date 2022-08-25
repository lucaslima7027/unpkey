import subprocess
import os

# Finds Office keys and return a list 
def get_keys():
   keys = []

   #  Run a new subprocess and capture the output
   p = subprocess.run('Cscript ospp.vbs /dstatus', stdout=subprocess.PIPE)    #  This command return the details of Office instaled including the key

   output = p.stdout                                  #  Captures the output
   output_decoded = output.decode('ISO-8859-1')       #  Decodes the output to one string
   output_splitted = output_decoded.split()           #  Splits the string in a list of strings

   #  Loop through list of strings to locate the Office keys
   for i in range(len(output_splitted)):
      if output_splitted[i] == "key:":                               #  Finds the string "key:" the next item on list is the Office key
         next_index = i + 1
         keys.append(output_splitted[next_index])                    #  Adds the Office key to a list
      i = i + 1
   
   return keys


#  Uninstall each Office key
def uninstall_keys():
   keys_to_uninstall = get_keys()            # Call the function to locate the keys

   if keys_to_uninstall:
      print(" Uninstalling Office Keys...")
      #  Loop through list and generates the command for Uninstall each Office key
      for i in range(len(keys_to_uninstall)):
         uninstall_command = "cscript ospp.vbs /unpkey:" + keys_to_uninstall[i]           # Attach Office key in command

         #  Run a new subprocess and capture the output
         p = subprocess.run(uninstall_command, stdout=subprocess.PIPE)              # This command Uninstall the Office key
         uninstall_output = p.stdout                                                # Capture the output
         uninstall_output_decoded = uninstall_output.decode('ISO-8859-1')           # Decode the output to string
         print("\n\n", uninstall_output_decoded)                                    # Print the Uninstall message 
         i = i + 1
      

   else:
      print(" No keys to Uninstall")
      

def main():
   
   print("\n-----PRODUCT KEY UNINSTALLER-----")
   #  List with directories
   paths = [r"C:\Program Files (x86)\Microsoft Office\Office16", r"C:\Program Files\Microsoft Office\Office16", r"C:\Program Files (x86)\Microsoft Office\Office15",r"C:\Program Files\Microsoft Office\Office15"]

   #  Loop through list and check if directory exists
   for i in range(len(paths)):
      print("\n Searching directory...")
      current_dir = paths[i]

      try:
         os.chdir(current_dir)
         print(" Directory found - " + paths[i])
         uninstall_keys()

      except:
         print(" Directory not found - " + paths[i])
   
   os.system(' pause')

main()


