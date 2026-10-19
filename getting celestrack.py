import argparse
import os
import pandas as pd
import requests


def fetch_and_convert_celestrak(
    url=None, group_name=None, output_path="output.xlsx"
):
  """Fetch orbital data (CSV/OMM or TLE format) from CelesTrak,

  parse it into a pandas DataFrame, and save it as an Excel file.
  """
  # Determine target URL based on input
  if group_name:
    url = f"https://celestrak.org/NORAD/elements/gp.php?GROUP={group_name}&FORMAT=csv"
  elif not url:
    raise ValueError("Either url or group_name must be provided.")

  print(f"Fetching data from: {url}")

  # Fetch data from CelesTrak
  response = requests.get(url, timeout=30)
  response.raise_for_status()

  # Check if data is CSV/OMM format or plain TLE
  content = response.text.strip()
  if content.startswith("OBJECT_NAME") or "," in content.splitlines()[0]:
    # Parse CSV format using pandas
    from io import StringIO

    df = pd.read_csv(StringIO(content))
  else:
    # If raw TLE text format, you can handle line parsing here
    raise NotImplementedError(
        "Raw TLE parsing format requires custom handling; please use"
        " FORMAT=csv."
    )

  # Ensure target directory exists
  output_dir = os.path.dirname(output_path)
  if output_dir and not os.path.exists(output_dir):
    os.makedirs(output_dir, exist_ok=True)

  # Save DataFrame to an Excel file with professional formatting
  with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Orbital Data", index=False)

  print(f"Successfully saved {len(df)} records to: {output_path}")
  return df


if __name__ == "__main__":
  # Command line argument parser
  parser = argparse.ArgumentParser(
      description="Download CelesTrak orbital data and save to Excel."
  )
  parser.add_argument(
      "--url", type=str, help="Direct CelesTrak GP data URL (FORMAT=csv)"
  )
  parser.add_argument(
      "--group", type=str, help="CelesTrak group name (e.g., irridium, stations)"
  )
  parser.add_argument(
      "-o",
      "--output",
      type=str,
      default="./orbital_data.xlsx",
      help="Target path and filename for output Excel",
  )

  args = parser.parse_args()
  fetch_and_convert_celestrak(
      url=args.url, group_name=args.group, output_path=args.output
  )
