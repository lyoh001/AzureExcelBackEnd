import asyncio
import base64
import functools
import os
import tempfile
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import aiofiles
import aiohttp
import numpy as np
import pandas as pd
import uvicorn
from dotenv import find_dotenv, load_dotenv
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook

load_dotenv(find_dotenv())
app = FastAPI()
id, origins = "", [
    "http://localhost:3000",
    os.environ["ORIGIN_0"],
    os.environ["ORIGIN_1"],
    os.environ["ORIGIN_2"],
]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["*"],
)
downloaded_files_path = "/tmp/downloaded_files.txt"
patterns = [
    ["CyberArk%20and%20DigiCert", "SOC"],
    ["Security%20Tools", "SOC"],
    ["SOC%202%20Services", "SOC"],
    ["Windows", "SOC 2 - Windows Privileged User Access"],
    ["Windows", "3402 - Windows Privileged User Access"],
    ["Windows", "3150 - Windows Privileged User Access"],
]


def get_id(request) -> str:
    id = (
        pn if (pn := request.headers.get("X-MS-CLIENT-PRINCIPAL-NAME")) else "anonymous"
    )
    return id.split("@")[0].replace(".", " ") if "@" in id else id


def get_api_headers_decorator(func):
    @functools.wraps(func)
    async def wrapper(session, *args, **kwargs):
        return {
            "Authorization": (
                f"Basic {base64.b64encode(bytes(os.environ[args[0]], 'utf-8')).decode('utf-8')}"
                if "PAT" in args[0]
                else f"Bearer {os.environ[args[0]] if 'EA' in args[0] else await func(session, *args, **kwargs)}"
            ),
            "Content-Type": (
                "application/json-patch+json"
                if "PAT" in args[0]
                else "application/json"
            ),
        }

    return wrapper


@get_api_headers_decorator
async def get_api_headers(session, *args, **kwargs):
    oauth2_headers = {"Content-Type": "application/x-www-form-urlencoded"}
    oauth2_body = {
        "client_id": os.environ[args[0]],
        "client_secret": os.environ[args[1]],
        "grant_type": "client_credentials",
        "scope" if "GRAPH" in args[0] else "resource": args[2],
    }
    async with session.post(
        url=args[3], headers=oauth2_headers, data=oauth2_body
    ) as resp:
        return (await resp.json())["access_token"]


async def fetch_data(session, url, headers):
    async with session.get(url=url, headers=headers) as resp:
        return await resp.json()


async def download_file(session, file_name, download_url):
    async with session.get(url=download_url) as resp:
        temp_file_path = os.path.join(
            tempfile.gettempdir(),
            file_name,
        )
        async with aiofiles.open(temp_file_path, "wb") as temp_file:
            await temp_file.write(await resp.content.read())
        return temp_file_path


async def download_file_async(
    session, drive_id, month, folder_name, file_pattern, graph_api_headers, url
):
    folder_month_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{url}/{month}/{folder_name}:/children"
    folder_month_data = await fetch_data(session, folder_month_url, graph_api_headers)
    file_name = next(
        f["name"] for f in folder_month_data["value"] if file_pattern in f["name"]
    )
    file_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{url}/{month}/{folder_name}/Test - {file_name}?select=id,@microsoft.graph.downloadUrl"
    file_data = await fetch_data(session, file_url, graph_api_headers)
    temp_file_path = await download_file(
        session, file_name, file_data["@microsoft.graph.downloadUrl"]
    )
    return temp_file_path


def save_downloaded_files_to_file(downloaded_files):
    with open(downloaded_files_path, "w") as file:
        file.write("\n".join(downloaded_files))


def load_downloaded_files_from_file():
    if os.path.exists(downloaded_files_path):
        with open(downloaded_files_path, "r") as file:
            return file.read().splitlines()
    return None


async def process_sheet_async(
    file_index,
    file_path_current,
    file_path_previous,
    sheet_pattern,
    column_indices,
    reviewer_name,
    remove_lastname,
):
    def process_sheet_sync():
        sheet_index_previous = [
            index
            for index, name in enumerate(pd.ExcelFile(file_path_previous).sheet_names)
            if sheet_pattern in name.lower()
        ][0]
        sheet_name_previous = pd.ExcelFile(file_path_previous).sheet_names[
            sheet_index_previous
        ]
        df_previous = pd.read_excel(file_path_previous, sheet_name=sheet_name_previous)
        df_previous.replace({np.nan: ""}, inplace=True)
        df_previous["ID"] = "ID"
        df_previous = df_previous[["ID"] + list(df_previous.columns[:-1])]
        df_previous = df_previous[
            df_previous.iloc[:, column_indices[0]]
            .str.lower()
            .str.contains(reviewer_name)
        ].iloc[
            :,
            column_indices[1:],
        ]
        df_previous.columns = [
            "ID",
            "Group",
            "Username",
            "Firstname",
            "Lastname",
            "LastApproval",
            "Remark",
        ]
        df_previous = df_previous.iloc[:, [1, 2, 3, 5]]

        sheet_index_current = [
            index
            for index, name in enumerate(pd.ExcelFile(file_path_current).sheet_names)
            if sheet_pattern in name.lower()
        ][0]
        sheet_name_current = pd.ExcelFile(file_path_current).sheet_names[
            sheet_index_current
        ]
        df_current = pd.read_excel(file_path_current, sheet_name=sheet_name_current)
        df_current.replace({np.nan: ""}, inplace=True)
        df_current["ID"] = f"{file_index}/{sheet_index_current}/" + (
            df_current.index + 2
        ).astype(str)
        df_current = df_current[["ID"] + list(df_current.columns[:-1])]
        df_current = df_current[
            df_current.iloc[:, column_indices[0]]
            .str.lower()
            .str.contains(reviewer_name)
        ].iloc[
            :,
            column_indices[1:],
        ]
        df_current["Filename"] = file_path_current.split("/")[2].split(".")[0]
        df_current["Sheetname"] = sheet_name_current
        df_current.columns = [
            "ID",
            "Group",
            "Username",
            "Firstname",
            "Lastname",
            "Approval",
            "Remark",
            "Filename",
            "Sheetname",
        ]
        df = pd.merge(
            left=df_current,
            right=df_previous,
            left_on=[
                df_current.columns[1],
                df_current.columns[2],
                df_current.columns[3],
            ],
            right_on=[
                df_previous.columns[0],
                df_previous.columns[1],
                df_previous.columns[2],
            ],
            how="left",
            indicator=False,
        )
        if remove_lastname:
            df["Lastname"] = ""
        df.replace({np.nan: ""}, inplace=True)
        return df

    with ThreadPoolExecutor() as executor:
        return await asyncio.get_event_loop().run_in_executor(
            executor, process_sheet_sync
        )


@app.get("/load")
async def load():
    async with aiohttp.ClientSession() as session:
        (graph_api_headers,) = await asyncio.gather(
            *(
                get_api_headers(session, *param)
                for param in [
                    [
                        "GRAPH_CLIENT_ID",
                        "GRAPH_CLIENT_SECRET",
                        "https://graph.microsoft.com/.default",
                        f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}/oauth2/v2.0/token",
                    ]
                ]
            )
        )
        drive_id, url = os.environ["DRIVE_ID"], os.environ["URL"]
        folder_url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{url}:/children"
        )
        folder_data = await fetch_data(session, folder_url, graph_api_headers)
        df = pd.DataFrame(folder_data["value"])
        df = df.sort_values(by="createdDateTime", ascending=False)
        download_coroutines = []
        for month in [
            df["name"].iloc[0].replace(" ", "%20"),
            df["name"].iloc[1].replace(" ", "%20"),
        ]:
            for pattern in patterns:
                coroutine = download_file_async(
                    session,
                    drive_id,
                    month,
                    pattern[0],
                    pattern[1],
                    graph_api_headers,
                    url,
                )
                download_coroutines.append(coroutine)

        downloaded_files = await asyncio.gather(*download_coroutines)
        save_downloaded_files_to_file(downloaded_files)

    processing_coroutines = [
        process_sheet_async(
            file_index,
            downloaded_files[file_index],
            downloaded_files[file_index + len(patterns)],
            sheet_pattern,
            column_indices,
            "jamero",
            remove_lastname,
        )
        for file_index, sheet_pattern, column_indices, remove_lastname in [
            # Month UAR - SOC 2 - CyberArk Privileged Users Confirmation (CyberArk)
            (
                0,
                "cyberark",
                [5, 0, 1, 2, 3, 4, 7, 9],
                False,
            ),
            # Month - SOC 2 - Security Tools Privileged User Access Confirmation (Cylance)
            (
                1,
                "cylance",
                [12, 0, 5, 10, 3, 4, 14, 16],
                False,
            ),
            # Month - SOC 2 - Security Tools Privileged User Access Confirmation (PKI Server Review)
            (
                1,
                "pki",
                [13, 0, 1, 3, 4, 5, 15, 17],
                False,
            ),
            # Month - UAR-SOC 2 Services - Access Confirmation (GO Desktop 365-SCCM)
            (
                2,
                "go desktop 365-sccm",
                [17, 0, 1, 4, 5, 8, 19, 21],
                False,
            ),
            # Month - UAR-SOC 2 Services - Access Confirmation (Go Office 365 additional groups)
            (
                2,
                "go office 365 additional groups",
                [4, 0, 1, 2, 2, 2, 3, 6],
                True,
            ),
            # Month - UAR-SOC 2 Services - Access Confirmation (Go Office 365)
            (
                2,
                "go office",
                [11, 0, 1, 3, 2, 2, 13, 15],
                True,
            ),
            # Month UAR - SOC 2 - Windows Privileged User Access Confirmation (GSP- INTERNAL  AD)
            (
                3,
                "internal ad acc",
                [24, 0, 3, 5, 6, 7, 26, 28],
                False,
            ),
            # Month UAR - SOC 2 - Windows Privileged User Access Confirmation (GSP- DOI  AD)
            (
                3,
                "doi ad acc",
                [22, 0, 3, 5, 6, 7, 23, 25],
                False,
            ),
            # Month UAR - SOC 2 - Windows Privileged User Access Confirmation (GSP-Workgroup Local Acc)
            (
                3,
                "workgroup local acc",
                [18, 0, 5, 7, 8, 8, 19, 21],
                True,
            ),
            # Month UAR - SOC 2 - Windows Privileged User Access Confirmation (DHHS-SERVICE AD)
            (
                3,
                "service ad acc",
                [22, 0, 3, 5, 6, 7, 23, 28],
                False,
            ),
            # Month UAR - SOC 2 - Windows Privileged User Access Confirmation (DJCS DOJVIC AD)
            (
                3,
                "dojvic  ad acc",
                [21, 0, 3, 5, 6, 7, 22, 27],
                False,
            ),
            # Month UAR - SOC 2 - Windows Privileged User Access Confirmation (PERIMETER AD)
            (
                3,
                "perimeter ad acc",
                [21, 0, 3, 5, 6, 7, 22, 27],
                False,
            ),
            # Month UAR - SOC 2 - Windows Privileged User Access Confirmation (CA Local)
            (
                3,
                "ca local acc",
                [15, 0, 5, 6, 2, 7, 16, 18],
                False,
            ),
            # Month UAR - SOC 2 - Windows Privileged User Access Confirmation (CA AD)
            (
                3,
                "ca ad acc",
                [22, 0, 3, 5, 6, 7, 23, 25],
                False,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (GSP-INTERNAL  Local)
            (
                4,
                "internal local acc",
                [15, 0, 5, 6, 7, 7, 17, 19],
                True,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (GSP- INTERNAL  AD)
            (
                4,
                "internal ad acc",
                [26, 0, 3, 5, 6, 7, 28, 30],
                False,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (GSP- DOI  AD)
            (
                4,
                "doi ad acc",
                [21, 0, 3, 5, 6, 7, 22, 24],
                False,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (GSP- DOI  Local)
            (
                4,
                "doi local acc",
                [16, 0, 5, 6, 7, 7, 17, 19],
                True,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (GSP-Workgroup Local Acc)
            (
                4,
                "workgroup local acc",
                [18, 0, 5, 7, 8, 8, 19, 21],
                True,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (DHHS-SERVICE AD)
            (
                4,
                "service ad acc",
                [21, 0, 3, 5, 6, 7, 22, 24],
                False,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (DHHS MGT Local)
            (
                4,
                "mgt local acc",
                [15, 0, 5, 6, 4, 7, 16, 18],
                False,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (DHHS MGT AD)
            (
                4,
                "mgt ad acc",
                [21, 0, 3, 5, 6, 7, 22, 24],
                False,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (DJCS DOJVIC AD)
            (
                4,
                "dojvic ad acc",
                [21, 0, 3, 5, 6, 7, 22, 24],
                False,
            ),
            # Month - ASAE 3402 - Windows Privileged User Access Confirmation (PERIMETER AD)
            (
                4,
                "perimeter ad acc",
                [21, 0, 3, 5, 6, 7, 22, 24],
                False,
            ),
            # Month UAR - 3150 - Windows Privileged User Access Confirmation (GSP- INTERNAL  AD)
            (
                5,
                "internal ad acc",
                [24, 0, 3, 5, 6, 7, 26, 27],
                False,
            ),
            # Month UAR - 3150 - Windows Privileged User Access Confirmation (GSP-Workgroup Local Acc)
            (
                5,
                "workgroup local acc",
                [18, 0, 5, 7, 4, 8, 19, 21],
                False,
            ),
            # Month UAR - 3150 - Windows Privileged User Access Confirmation (DHHS-SERVICE Local Acc)
            (
                5,
                "service local acc",
                [16, 0, 5, 6, 6, 6, 17, 19],
                True,
            ),
            # Month UAR - 3150 - Windows Privileged User Access Confirmation (DHHS-SERVICE AD)
            (
                5,
                "service ad acc",
                [22, 0, 3, 5, 6, 7, 23, 28],
                False,
            ),
            # Month UAR - 3150 - Windows Privileged User Access Confirmation (DHHS MGT Local)
            (
                5,
                "mgt local acc",
                [15, 0, 5, 6, 4, 7, 16, 18],
                False,
            ),
            # Month UAR - 3150 - Windows Privileged User Access Confirmation (DHHS MGT AD)
            (
                5,
                "mgt ad acc",
                [21, 0, 3, 5, 6, 7, 22, 27],
                False,
            ),
        ]
        if file_index < len(patterns)
    ]
    processed_data = await asyncio.gather(*processing_coroutines)
    df = pd.concat(processed_data, axis=0, ignore_index=True)
    df.loc[
        df["Firstname"]
        .str.lower()
        .str.contains(
            r"\b(?:{})\b".format(
                "|".join(
                    [
                        "al",
                        "alan",
                        "candi",
                        "candido",
                        "chandra",
                        "dhruv",
                        "frank",
                        "glenn",
                        "john",
                        "miguel",
                        "mino",
                        "prashanth",
                        "ralph",
                        "rod",
                        "sofya",
                        "suhail",
                        "sushma",
                        "vikrant",
                        "zachary",
                    ]
                )
            )
        )
        & (df["Approval"] == ""),
        "Approval",
    ] = "Y"
    df.sort_values(
        by=["Approval", "Sheetname", "Group", "Firstname"],
        ascending=[True, True, True, True],
        inplace=True,
    )
    return df.to_dict(orient="records")


async def modify_file(file_index, file_path, data, id, time):
    workbook = load_workbook(file_path)
    for cell_id, approval in data["data"]["approvals"].items():
        index, sheet_index, row_number = map(int, cell_id.split("/"))
        if index == file_index:
            sheet = workbook.worksheets[sheet_index]
            row = sheet[row_number]
            if file_index == 0 and "cyberark" in sheet.title.lower():
                row[5].value = time
                row[6].value = approval
                row[9].value = id
                row[10].value = time
            elif file_index == 1 and "cylance" in sheet.title.lower():
                row[12].value = time
                row[13].value = approval
                row[16].value = id
                row[17].value = time
            elif file_index == 1 and "pki" in sheet.title.lower():
                row[13].value = time
                row[14].value = approval
                row[17].value = id
                row[18].value = time
            elif file_index == 2 and "go desktop 365-sccm" in sheet.title.lower():
                row[17].value = time
                row[18].value = approval
                row[21].value = id
                row[22].value = time
            elif (
                file_index == 2
                and "go office 365 additional groups" in sheet.title.lower()
            ):
                row[4].value = time
                row[2].value = approval
                row[6].value = id
                row[7].value = time
            elif file_index == 2 and "go office" in sheet.title.lower():
                row[11].value = time
                row[12].value = approval
                row[15].value = id
                row[16].value = time
            elif file_index == 3 and "internal ad acc" in sheet.title.lower():
                row[24].value = time
                row[25].value = approval
                row[28].value = id
                row[29].value = time
            elif file_index == 3 and "doi ad acc" in sheet.title.lower():
                row[20].value = time
                row[22].value = approval
                row[25].value = id
                row[26].value = time
            elif file_index == 3 and "workgroup local acc" in sheet.title.lower():
                row[16].value = time
                row[18].value = approval
                row[21].value = id
                row[22].value = time
            elif file_index == 3 and "service ad acc" in sheet.title.lower():
                row[20].value = time
                row[22].value = approval
                row[25].value = id
                row[26].value = time
            elif file_index == 3 and "dojvic  ad acc" in sheet.title.lower():
                row[19].value = time
                row[21].value = approval
                row[24].value = id
                row[25].value = time
            elif file_index == 3 and "perimeter ad acc" in sheet.title.lower():
                row[19].value = time
                row[21].value = approval
                row[24].value = id
                row[25].value = time
            elif file_index == 3 and "ca local acc" in sheet.title.lower():
                row[13].value = time
                row[15].value = approval
                row[18].value = id
                row[19].value = time
            elif file_index == 3 and "ca ad acc" in sheet.title.lower():
                row[20].value = time
                row[22].value = approval
                row[25].value = id
                row[26].value = time
            elif file_index == 4 and "internal local acc" in sheet.title.lower():
                row[15].value = time
                row[16].value = approval
                row[19].value = id
                row[20].value = time
            elif file_index == 4 and "internal ad acc" in sheet.title.lower():
                row[26].value = time
                row[27].value = approval
                row[30].value = id
                row[31].value = time
            elif file_index == 4 and "doi ad acc" in sheet.title.lower():
                row[19].value = time
                row[21].value = approval
                row[24].value = id
                row[25].value = time
            elif file_index == 4 and "doi local acc" in sheet.title.lower():
                row[14].value = time
                row[16].value = approval
                row[19].value = id
                row[20].value = time
            elif file_index == 4 and "workgroup local acc" in sheet.title.lower():
                row[16].value = time
                row[18].value = approval
                row[21].value = id
                row[22].value = time
            elif file_index == 4 and "service ad acc" in sheet.title.lower():
                row[19].value = time
                row[21].value = approval
                row[24].value = id
                row[25].value = time
            elif file_index == 4 and "mgt local acc" in sheet.title.lower():
                row[13].value = time
                row[15].value = approval
                row[18].value = id
                row[19].value = time
            elif file_index == 4 and "mgt ad acc" in sheet.title.lower():
                row[19].value = time
                row[21].value = approval
                row[24].value = id
                row[25].value = time
            elif file_index == 4 and "dojvic ad acc" in sheet.title.lower():
                row[19].value = time
                row[21].value = approval
                row[24].value = id
                row[25].value = time
            elif file_index == 4 and "perimeter ad acc" in sheet.title.lower():
                row[19].value = time
                row[21].value = approval
                row[24].value = id
                row[25].value = time
            elif file_index == 5 and "internal ad acc" in sheet.title.lower():
                row[24].value = time
                row[25].value = approval
                row[28].value = id
                row[29].value = time
            elif file_index == 5 and "workgroup local acc" in sheet.title.lower():
                row[16].value = time
                row[18].value = approval
                row[21].value = id
                row[22].value = time
            elif file_index == 5 and "service local acc" in sheet.title.lower():
                row[14].value = time
                row[16].value = approval
                row[19].value = id
                row[20].value = time
            elif file_index == 5 and "service ad acc" in sheet.title.lower():
                row[20].value = time
                row[22].value = approval
                row[25].value = id
                row[26].value = time
            elif file_index == 5 and "mgt local acc" in sheet.title.lower():
                row[13].value = time
                row[15].value = approval
                row[18].value = id
                row[19].value = time
            elif file_index == 5 and "mgt ad acc" in sheet.title.lower():
                row[19].value = time
                row[21].value = approval
                row[24].value = id
                row[25].value = time

    for cell_id, remark in data["data"]["remarks"].items():
        index, sheet_index, row_number = map(int, cell_id.split("/"))
        if index == file_index:
            sheet = workbook.worksheets[sheet_index]
            row = sheet[row_number]
            if file_index == 0 and "cyberark" in sheet.title.lower():
                row[8].value = remark
            elif file_index == 1 and "cylance" in sheet.title.lower():
                row[15].value = remark
            elif file_index == 1 and "pki" in sheet.title.lower():
                row[16].value = remark
            elif file_index == 2 and "go desktop 365-sccm" in sheet.title.lower():
                row[20].value = remark
            elif (
                file_index == 2
                and "go office 365 additional groups" in sheet.title.lower()
            ):
                row[5].value = remark
            elif file_index == 2 and "go office" in sheet.title.lower():
                row[14].value = remark
            elif file_index == 3 and "internal ad acc" in sheet.title.lower():
                row[27].value = remark
            elif file_index == 3 and "doi ad acc" in sheet.title.lower():
                row[24].value = remark
            elif file_index == 3 and "workgroup local acc" in sheet.title.lower():
                row[20].value = remark
            elif file_index == 3 and "service ad acc" in sheet.title.lower():
                row[27].value = remark
            elif file_index == 3 and "dojvic  ad acc" in sheet.title.lower():
                row[26].value = remark
            elif file_index == 3 and "perimeter ad acc" in sheet.title.lower():
                row[26].value = remark
            elif file_index == 3 and "ca local acc" in sheet.title.lower():
                row[17].value = remark
            elif file_index == 3 and "ca ad acc" in sheet.title.lower():
                row[24].value = remark
            elif file_index == 4 and "internal local acc" in sheet.title.lower():
                row[18].value = remark
            elif file_index == 4 and "internal ad acc" in sheet.title.lower():
                row[29].value = remark
            elif file_index == 4 and "doi ad acc" in sheet.title.lower():
                row[23].value = remark
            elif file_index == 4 and "doi local acc" in sheet.title.lower():
                row[18].value = remark
            elif file_index == 4 and "workgroup local acc" in sheet.title.lower():
                row[20].value = remark
            elif file_index == 4 and "service ad acc" in sheet.title.lower():
                row[23].value = remark
            elif file_index == 4 and "mgt local acc" in sheet.title.lower():
                row[17].value = remark
            elif file_index == 4 and "mgt ad acc" in sheet.title.lower():
                row[23].value = remark
            elif file_index == 4 and "dojvic ad acc" in sheet.title.lower():
                row[23].value = remark
            elif file_index == 4 and "perimeter ad acc" in sheet.title.lower():
                row[23].value = remark
            elif file_index == 5 and "internal ad acc" in sheet.title.lower():
                row[27].value = remark
            elif file_index == 5 and "workgroup local acc" in sheet.title.lower():
                row[20].value = remark
            elif file_index == 5 and "service local acc" in sheet.title.lower():
                row[18].value = remark
            elif file_index == 5 and "service ad acc" in sheet.title.lower():
                row[27].value = remark
            elif file_index == 5 and "mgt local acc" in sheet.title.lower():
                row[17].value = remark
            elif file_index == 5 and "mgt ad acc" in sheet.title.lower():
                row[26].value = remark

    workbook.save(file_path)
    workbook.close()


@app.post("/update")
async def update_data(request: Request, data: dict):
    async with aiohttp.ClientSession() as session:
        (graph_api_headers,) = await asyncio.gather(
            *(
                get_api_headers(session, *param)
                for param in [
                    [
                        "GRAPH_CLIENT_ID",
                        "GRAPH_CLIENT_SECRET",
                        "https://graph.microsoft.com/.default",
                        f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}/oauth2/v2.0/token",
                    ]
                ]
            )
        )
        drive_id, url = os.environ["DRIVE_ID"], os.environ["URL"]
        folder_url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{url}:/children"
        )
        folder_data = await fetch_data(session, folder_url, graph_api_headers)
        df = pd.DataFrame(folder_data["value"])
        df = df.sort_values(by="createdDateTime", ascending=False)
        download_coroutines = []
        for month in [
            df["name"].iloc[0].replace(" ", "%20"),
            df["name"].iloc[1].replace(" ", "%20"),
        ]:
            for pattern in patterns:
                coroutine = download_file_async(
                    session,
                    drive_id,
                    month,
                    pattern[0],
                    pattern[1],
                    graph_api_headers,
                    url,
                )
                download_coroutines.append(coroutine)

        downloaded_files = await asyncio.gather(*download_coroutines)
        save_downloaded_files_to_file(downloaded_files)
        id = data["data"]["userInfo"].split("@")[0].replace(".", " ").title()
        time = datetime.now().strftime("%d/%m/%Y")
        modification_tasks = []
        for file_index, file_path in enumerate(
            downloaded_files[: len(downloaded_files) // 2]
        ):
            task = modify_file(file_index, file_path, data, id, time)
            modification_tasks.append(task)
        await asyncio.gather(*modification_tasks)

        month = df["name"].iloc[0].replace(" ", "%20")
        upload_tasks = []
        for file_index, file_path in enumerate(
            downloaded_files[: len(downloaded_files) // 2]
        ):
            file_name = os.path.basename(file_path)
            folder_name = patterns[file_index][0]
            async with aiofiles.open(file_path, "rb") as f:
                file_content = await f.read()
                upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{url}/{month}/{folder_name}/Test - {file_name}:/content"
                task = session.put(
                    url=upload_url,
                    headers=graph_api_headers,
                    data=file_content,
                )
                upload_tasks.append(task)
        upload_responses = await asyncio.gather(*upload_tasks)
        for resp in upload_responses:
            if resp.status != 200:
                return {"message": "Error occurred while uploading the file"}

    return {"message": "Data updated and uploaded successfully"}


if __name__ == "__main__":
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)
