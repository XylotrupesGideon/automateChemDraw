import os
import win32com.client as win32
import argparse


def is_valid_format(output_format):
    # List of valid ChemDraw output formats
    valid_formats = [
        "cdx",
        "cdxml",
        "chm",
        "cds",
        "ct",
        "cml",
        "bmp",
        "tif",
        "gif",
        "skc",
        "jdx",
        "jpg",
        "jpeg",
        "sdf",
        "mol",
        "rdf",
        "spc",
        "wmf",
        "emf",
        "ctp",
        "cts",
        "pict",
        "pict4x",
        "eps",
        "svg",
        "3mf",
        "cif",
        "mmcif",
        "png",
    ]

    return output_format.lower() in valid_formats


def convert_cdxml_to_formats(input_folder, output_formats):
    # Create an instance of the ChemDraw application
    cd_app = win32.Dispatch("ChemDraw.Application")
    print("Starting")

    # Validate output formats
    invalid_formats = [
        format for format in output_formats if not is_valid_format(format)
    ]
    if invalid_formats:
        print(f"Error: Invalid output formats specified: {', '.join(invalid_formats)}")
        return

    # Traverse the input folder and its subfolders
    for foldername, subfolders, filenames in os.walk(input_folder):
        for filename in filenames:
            if filename.lower().endswith(".cdx"):
                # Construct the full path to the CDXML file
                cdxml_path = os.path.abspath(os.path.join(foldername, filename))
                print(cdxml_path)

                try:
                    # Open the CDXML file in ChemDraw
                    cd_doc = cd_app.Documents.Open(cdxml_path)
                    cd_app.Visible = True

                    # Process each specified output format
                    for output_format in output_formats:
                        # Save the file in the specified format
                        output_path = (
                            os.path.splitext(cdxml_path)[0]
                            + f".{output_format.lower()}"
                        )
                        cd_doc.SaveAs(output_path, Format=output_format)

                    # Close the ChemDraw document
                    cd_doc.Close()
                except Exception as e:
                    print(f"Error processing {cdxml_path}: {e}")

    # Close ChemDraw
    cd_app.Quit()


def main():
    parser = argparse.ArgumentParser(
        description="Convert CDXML files to specified formats."
    )
    parser.add_argument(
        "input_folder", help="Path to the folder containing CDXML files."
    )
    parser.add_argument(
        "output_formats",
        nargs="+",
        help="List of output formats (e.g., cdx cdxml pdf).",
    )

    args = parser.parse_args()

    convert_cdxml_to_formats(args.input_folder, args.output_formats)


if __name__ == "__main__":
    main()
