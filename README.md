# WMI-Power

## Purpose

This Powershell script is a utility script that helped me maintain an acurate inventory of 5,000 computers.
It was written for a specific environment and is meant to be a guide to those looking to do something similar. 

## How it works

It iterates through a list of hostnames via the ComputerList.csv. It first attempts to ping the computer if the computer
is unreachable it will output a csv with the devices it could not connect to. As for the pingable computers it gathers
data from them and outputs it to a csv that is in a time stamped directory. The output csv is in a specific arrangement
which worked for the environment I was in and the previous inventory I was updating.  

## Mods

The output is easily changed by changing the `getObject` function. There is a PSObject `$InventoryItem` that is directly output to the inventory csv.  
