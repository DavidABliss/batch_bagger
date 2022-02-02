# batch_bagger
Create many bags at once, using a bag-info template and a CSV or XLSX spreadsheet to substitute bag-level text as needed

Usage: 

<code>python batch_bagger.py --directory <path_to_bags_directory> --baginfo <path_to_bag-info_template_txt_file> --csv <path_to_bag-info_CSV_spreadsheet> --verbose</code>

This script compiles folders in a target directory into individual bags. Users supply a target directory, a bag-info.txt template file, and a CSV or XLSX spreadsheet containing bag-level information. The script reads the CSV or XLSX spreadsheet, matches each entry to a bag in the target directory, and builds a bag-info file using the bag-info template file, supplemented with any bag-specific information provided in the spreadsheet.

At a minimum, the CSV or XLSX spreadsheet must have a column headers row, and the first column must list the names of the folders to be bagged within the target directory. In this case, the only difference between the bag-info template and the bag-info file in each bag will be the addition of a UUID External-Identifier

To supplement the bag-info template file with bag-specific information, enclose placeholder text in the bag-info template within [[double brackets]]. Create a column with a header matching that text (minus the brackets) in the CSV or XLSX spreadsheet, and put bag-specific replacement text in that column. The script will match the placeholder text to the column header and insert the corresponding value into the output bag-info.txt file.

Values in the bag-info template file can be split across multiple lines. The script will combine these values into a single line for the output bag-info.txt file.

Fields in the bag-info template file can repeat. bagit.py does not allow fields to be repeated, so each value will be combind into a single line, separated by pipes.

The script assigns a UUID External-Identifier to each bag, but allows for existing External-Identifiers to be provided alongside this.

Users can add the <code>-v --verbose</code> switch to print the contents of the bag-info file as each bag is created

All bags are created in place.

Finally, the script outputs a CSV list of bags that were created, alongside the UUID assigned to each. This is meant for recordkeeping, but can be modified or removed as needed.

This script is written for the UT Libraries' local bag-info spec, and will not recognize any terms not found in fieldsList as labels for the output bag-info.txt files. The script can be made to work with other bag-info specs by modifying that list.

Users can add the <code>-u --unpack</code> switch to "unbag" a series of folders. If a spreadsheet is provided, bags listed in the first column will be unpacked within the target directory (cwd by default). If no spreadsheet is provided, the script will test each folder in the target directory and skip any folders not containing a bag-info.txt file. Because the unpack function will invalidate bags in the target directory, the script displays a warning message and asks the user to confirm that they wish to proceed.
