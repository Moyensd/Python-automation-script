Attached is the python script as promised. It's in txt format so you can open it and take a look at it as well. lines after "#" are the comments about the functionality.

What this script does is copy the excel worksheet table from our audits and saves it to input it into denticon. It is missing cursor positioning data entirely and I only put some random numbers I had while debugging the script so right now it won't work until I change it live on the work computer. It is designed to only do one chart at a time.

The script uses the windows to web browser GUI to operate and execute the automation so it's fairly low level automation and cannot be run in the background. That might not be possible anyway since denticon is owned by planetdds and the API is not public.

Some possible future improvements could still be made however.
openpyxl module can be used to extract the data from the xlsx file directly and avoid opening it from onedrive.
playwright module can also help account for web browser loading and time the script out if denticon is being slow.
With some development it is likely the entire process of moving the pending tx from excel to denticon can be automated.
Python scripts can also be run through microsoft power automate desktop with their cloud service so if approved and tested any RCM computer can run it with microsoft power automate without python installed.

If you have any questions please ask.

Thank you for your time reading this.
