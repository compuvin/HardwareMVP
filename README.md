# HardwareMVP
Modular Hardware Inventory solution that takes multiple inputs and compares the hardware that exists on all of them.

Requires MySQL 8.0
Script will prompt for server information and create the necessary tables.

HardwareMVP.vbs is, at its core, a scheduler. The frequency that it is run should be daily. The script inventories the Modules folder and adds/removes any modules that are in that folder keeping a modules table up-to-date in the database. It then starts the process for any modules that are scheduled to run on that day. It also runs any new ones that it finds. Each module can have its own frequency and that is defined by the PSRunInt variable in the module’s code. Once all modules have been run for the day, the data is compared to the Master List (specified by the PSML variable). Only one module can be the Master List (we use PDQ as our Master List). Any items removed from the Master List need to be fully removed from all of the modules’ lists before it is completely removed from the Master List. Modules have their own email alerts (even the Master List) but the HardwareMVP.vbs sends its own email alerts for items that are removed or remain to be removed from the Master List.

This project is still in beta. I can use all the help I can get if anyone is interested!

[![CodeFactor](https://www.codefactor.io/repository/github/compuvin/hardwaremvp/badge/main)](https://www.codefactor.io/repository/github/compuvin/hardwaremvp/overview/main)