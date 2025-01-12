import streamlit as st
import docx
from docx import Document
from docx.shared import Inches, Pt
import csv, wget, json
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.dml import MSO_THEME_COLOR_INDEX
from docx.enum.table import WD_ALIGN_VERTICAL
from operator import itemgetter
import numpy as np
from docx.oxml.shared import OxmlElement, qn
from docx.enum.text import WD_BREAK
from zipfile2 import ZipFile
import io

def find_positions_file(zipf):
    # for instance, find the first file that contains "positions" and ends with ".txt"
    for name in zipf.namelist():
        lower_name = name.lower()
        if "positions" in lower_name and lower_name.endswith(".txt"):
            return name
    # if we never return inside the loop, there's no match
    return None

# -----------------------------------------------------------------------------
# Streamlit Setup
# -----------------------------------------------------------------------------
st.title("Kumu File Generator with Download Buttons")
st.write("""
Upload a `.txps` file, parse it, and generate:
- A CSV showing persona permissions
- Kumu JSON output
- A CSV for Kumu connections
""")

# -----------------------------------------------------------------------------
# Constants & Globals
# -----------------------------------------------------------------------------
missing_person_url = "https://conducttr-a.akamaihd.net/fa/2006/06/16832/profile-trns-bg.png"

# We'll store final Kumu data globally here
kumu = {"elements": [], "connections": []}
do_the_lot = True  # You can adjust this logic as needed in your code

# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------
def create_connection(from_this, to_that, with_direction, ctype):
    connection = {
        "from": from_this,
        "to": to_that,
        "direction": with_direction,
        "type": ctype,
    }
    return connection

def create_kumu_element(
    id,
    label,
    handle,
    image,
    etype,
    faction,
    a3e,
    affiliation,
    bio,
    goals,
    permissions,
    mbfollowing,
    mbfollowers,
    gps
):
    ke = {
        "_id": id,
        "uid": id,
        "attributes": {},
        "label": label,
        "handle": handle,
        "image": image,
        "element type": etype,
        "faction": faction,
        "A3E": a3e,
        "Affiliation": affiliation,
        "bio": bio,
        "goals": goals,
        "permissions": permissions,
        "microblog_following": mbfollowing,
        "microblog_followers": mbfollowers,
        "location":gps
    }
    return ke

def return_name_and_handle_from_kumu(kumu_data, uid):
    for item in kumu_data["elements"]:
        if item["uid"] == uid:
            name = item["label"]
            handle = item["handle"]
            a3e = item["A3E"]
            return name, handle, a3e
    return "", "", ""

def export_kumu_to_csv(kumu_data, filename_for_download="hero_connections_i2.csv"):
    """
    Creates a CSV with Kumu 'connections' data.
    Then provides a download button for that CSV.
    """
    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    
    # Write header
    header_row = [
        "From",
        "Fname",
        "Fhandle",
        "Fa3e",
        "To",
        "Tname",
        "Thandle",
        "Ta3e",
        "Direction",
        "Type",
    ]
    writer.writerow(header_row)
    
    # Populate rows
    for row in kumu_data["connections"]:
        from_id = row["from"]
        to_id = row["to"]
        direction = row["direction"]
        ctype = row["type"]
        
        newrow = []
        # "From" columns
        newrow.append(from_id)
        fname, fhandle, fa3e = return_name_and_handle_from_kumu(kumu_data, from_id)
        newrow.append(fname)
        newrow.append(fhandle)
        newrow.append(fa3e)
        
        # "To" columns
        newrow.append(to_id)
        tname, thandle, ta3e = return_name_and_handle_from_kumu(kumu_data, to_id)
        newrow.append(tname)
        newrow.append(thandle)
        newrow.append(ta3e)
        
        # direction + type
        newrow.append(direction)
        newrow.append(ctype)
        writer.writerow(newrow)

    # Streamlit download button
    st.download_button(
        label="Download Connections CSV",
        data=output.getvalue(),
        file_name=filename_for_download,
        mime="text/csv",
    )
    st.write("Exported connections CSV to download.")

def check_if_missingv2(text):
    if text is None or text == "":
        return "MISSING"
    else:
        return text

def get_tier(tags_string):
    tiers = ["t1", "t2", "t3", "t4", "t5", "t6", "t7", "T1", "T2", "T3", "T4", "T5", "T6", "T7"]
    return_tier = 99
    for tier in tiers:
        if tier in tags_string:
            # The tier is always "t<number>" or "T<number>"
            return_tier = int(tier[1:])
            break
    return return_tier

def get_A3E(tags_string):
    a3e = ["Actor", "Adversary", "Audience", "Enemy"]
    for role in a3e:
        if role in tags_string:
            return role
    return ""

def get_affiliation(tags_string):
    nato = ["Unknown", "Friend", "Neutral", "Hostile"]
    for role in nato:
        if role in tags_string:
            return role
    return ""

def get_persona_record(personas, persona_id):
    for persona in personas:
        if persona["id"] == persona_id:
            return persona
    return {}

def main():

    
    # Step 1: Upload the .txps file
    st.subheader("Step 1: Upload TXPS File")
    uploaded_txps = st.file_uploader("Choose a .txps file", type=["txps"])
    if not uploaded_txps:
        st.warning("Please upload a TXPS file to proceed.")
        return

    # The original code references "positions id=2.txt" and "characters.txt"
    positions_file = "positions id=2.txt"
    characters_file = "characters.txt"

    # Step 2: Extract the needed files from the .txps (zip)
    with ZipFile(uploaded_txps, "r") as zipf:
        # Try to find a positions file
        positions_file = find_positions_file(zipf)
        if not positions_file:
            st.error("No valid 'positions' file found inside the ZIP.")
            return

        st.write(f"Found positions file: {positions_file}")

        with zipf.open(positions_file) as f:
            positions_data = json.load(f)
            
    # Build position_names from positions_data
    position_names = {}
    for item in positions_data["list"]:
        position_names[item["id"]] = item["name"]

    # Create top row for the permissions matrix
    toprow = ["Persona","Tier","A3E","Affiliation","RolePlayer"]
    for key in position_names:
        toprow.append(position_names[key])
    toprow.append("count")

    characters_file = "characters.txt"

    with ZipFile(uploaded_txps, "r") as zipf:
        # 1. Check if characters_file exists in the zip
        if characters_file not in zipf.namelist():
            st.error(f"No {characters_file} found inside the TXPS.")
            return
        
        # 2. Extract and load characters_data
        with zipf.open(characters_file) as f:
            characters_data = json.load(f)

    # Prepare arrays for storing permissions data
    if "list" in characters_data:
        data_list = characters_data["list"]
    elif "items" in characters_data:
        data_list = characters_data["items"]
    else:
        st.error("No 'list' or 'items' found in characters.txt JSON. Please verify the TXPS format.")
        return

    st.write(data_list)
    
    #data_list = characters_data["list"]
    permissions_count = {}
    permissions_str = {}

    # Our final matrix (row_count x col_count)
    row_count = len(data_list) + 1
    col_count = len(toprow)
    #array = np.full((row_count, col_count), " ")
    array = np.full((row_count, col_count), "", dtype=object)
    
    # Set header row
    array[0] = toprow

    # Step 3: Populate the matrix from the scenario data
    row_index = 1
    for persona in data_list:
        st.write(persona)
        uid = persona["uid"]
        permissions_str[uid] = " "
        permissions_count[uid] = 0

        name = persona["name"].replace(",", "_")
        tags = persona["tags"]
        tier = get_tier(tags)
        a3e = get_A3E(tags)
        nato = get_affiliation(tags)
        roleplayer = persona["is_role_player"]

        # Place data into array
        array[row_index][0] = name
        array[row_index][1] = tier
        array[row_index][2] = a3e
        array[row_index][3] = nato
        array[row_index][4] = roleplayer

        # Parse permissions
        permissions = json.loads(persona["permissions"])
        count = 0
        if permissions.get("ids"):
            for pos in permissions["ids"]:
                try:
                    col_idx = toprow.index(position_names[pos])
                    array[row_index][col_idx] = "x"
                    count += 1
                    array[row_index][col_count - 1] = count
                    if count == 1:
                        permissions_str[uid] = position_names[pos] + ", "
                    else:
                        permissions_str[uid] += position_names[pos] + ", "
                    permissions_count[uid] = count
                except KeyError:
                    pass
        row_index += 1

    st.write("Persona permissions matrix has been created.")

    # Optional: Show preview of the matrix
    st.write("Preview (first 10 rows):")
    st.write(array[:10])

    # Step 4: Add a Download button for the persona permissions CSV
    output_csv = io.StringIO()
    csv_writer = csv.writer(output_csv, lineterminator="\n")
    for rowdata in array:
        csv_writer.writerow(rowdata)

    st.download_button(
        label="Download Persona Permissions CSV",
        data=output_csv.getvalue(),
        file_name="elements_i2.csv",
        mime="text/csv",
    )

    # Step 5: Build the Kumu data (elements + connections)
    persona_connections = []
    mbfollowed = {}
    kumu_list = []

    for persona in data_list:
        uid = persona["uid"]
        # Include this persona if it has permissions or is a roleplayer or do_the_lot
        if permissions_count[uid] > 0 or persona["is_role_player"] or do_the_lot:
            name = check_if_missingv2(persona["name"])
            handle = check_if_missingv2(persona["handle"])
            bio = check_if_missingv2(persona["bio"])
            image = check_if_missingv2(persona["image_url"])
            tags = persona["tags"]
            tier = get_tier(tags)
            a3e = get_A3E(tags)
            nato = get_affiliation(tags)

            # Microblog data
            twitter = json.loads(persona["microblog"])
            try:
                mbfollowing = len(twitter["following_ids"])
            except:
                mbfollowing = 0

            # Build the microblog "follow" connections
            if mbfollowing > 0:
                for following_id in twitter["following_ids"]:
                    p = get_persona_record(data_list, following_id)
                    if p:
                        # track how many times a persona is followed
                        if following_id in mbfollowed:
                            mbfollowed[following_id] += 1
                        else:
                            mbfollowed[following_id] = 1
                        c = create_connection(uid, p["uid"], "directed", "microblog")
                        persona_connections.append(c)

            # Build the Kumu element
            # Trim trailing comma if needed:
            persona_permissions = permissions_str[uid][:-2] if permissions_str[uid] != " " else ""
            new_kp = create_kumu_element(
                uid,
                name,
                handle,
                image,
                "Person",
                "",
                a3e,
                nato,
                bio,
                persona.get("goals", ""),  # goals might be a string or JSON array
                persona_permissions,
                mbfollowing,
                0,  # will update below
                persona.get("gps", "")
            )
            kumu_list.append(new_kp)

    # Update follower counts
    for elem in kumu_list:
        uid = elem["uid"]
        elem["mbfollowers"] = mbfollowed.get(uid, 0)

    # Final Kumu data
    kumu["elements"] = kumu_list
    kumu["connections"] = persona_connections

    # Step 6: Provide a download button for the Kumu JSON
    kumu_json_data = json.dumps(kumu, indent=2)
    st.download_button(
        label="Download Kumu JSON",
        data=kumu_json_data,
        file_name="ws1.json",
        mime="application/json",
    )

    # Step 7: Export Kumu connections to CSV (with a download button)
    export_kumu_to_csv(kumu, "hero_connections_i2.csv")


# -----------------------------------------------------------------------------
# Run the Streamlit app
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    main()
