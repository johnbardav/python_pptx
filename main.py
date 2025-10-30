"""
Main script to generate the PowerPoint presentation
for the Bank App Analysis project.
"""

import sys
import os

try:
    from pptx import Presentation
    from pptx.util import Cm

    # Import all necessary functions from our 'masters' package
    from masters import create_base_slide, load_master_excels
except ImportError as e:
    print(f"Import Error: {e}")
    print("Please ensure 'python-pptx', 'pandas', and 'unidecode' are installed")
    print("and you are running this script from the project root directory.")
    sys.exit(1)
except Exception as e:
    print(f"An unexpected error occurred during imports: {e}")
    sys.exit(1)


def build_presentation(filename: str):
    """
    Creates and saves the complete PowerPoint presentation.
    """

    # --- 1. Load Data ---
    # Note: This loads from Excel. We will later change this
    # to load from the MySQL DB.
    print("\nLoading master Excel files...")
    (
        df_buyer,
        choices_buyer,
        df_bought,
        choices_bought,
    ) = load_master_excels()

    if df_buyer is None or df_bought is None:
        print("Could not load one or more Excel files. Aborting presentation build.")
        return

    print(
        f"Loaded {len(df_buyer)} Buyer Bank apps and {len(df_bought)} Bought Bank apps."
    )

    # --- 2. Create Presentation Object ---
    prs = Presentation()

    # Set presentation dimensions (Widescreen 16:9)
    prs.slide_width = Cm(33.87)
    prs.slide_height = Cm(19.05)

    print("\nCreating Widescreen (16:9) presentation...")

    # --- 3. Build Slides (Incremental) ---
    create_base_slide(
        prs=prs,
        title_text="Title of the First Slide",
        content_text="This is the content for the first slide of the project.",
    )

    create_base_slide(
        prs=prs,
        title_text="Second Incremental Slide",
        content_text="Here is different information.\n"
        "The layout is identical to the previous one.",
    )

    # --- 4. Save the File ---
    try:
        prs.save(filename)
        print(f"\nPresentation successfully saved as '{filename}'!")
    except PermissionError:
        print(f"\nError: Could not save file '{filename}'.")
        print("Please ensure the file is not already open in PowerPoint.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")


# --- Main execution block ---
if __name__ == "__main__":
    # Define the output directory
    output_dir = "outputs"

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created directory: {output_dir}")

    # Define the full output path
    output_filename = os.path.join(output_dir, "Bank_App_Analysis_output.pptx")

    build_presentation(output_filename)
