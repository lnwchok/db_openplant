import pyodbc
import pandas as pd

folder = r"D:\E\My_Dev\python\DB"

filename = f".\db\DB.accdb"
# filename = f"{folder}\db\DB.accdb"

# Connect to the database
conn = pyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;" % filename
)

# Create a cursor object
cursor = conn.cursor()

# SQL Statement
SQL_Get_oo7 = r"SELECT A.Name, A.MAT_ID, A.SIZE_ID, B.CAT FROM OO7 as A LEFT JOIN TBL_CAT as B ON A.MAT_ID = B.ID;"

# 0	No Need
# 1	Pipe
# 2	Flange
# 3	Fitting
# 4	Gasket
# 5	Bolt
# 6	Valve
# 7	Piping Accessory
# 8	Instrument

matl = {}
matl_sql = {}

# Iterate database to got list of OpenPlant Table
cursor.execute(SQL_Get_oo7)
for openplant_tbl in cursor.fetchall():
    if not openplant_tbl[3] in matl:
        matl[openplant_tbl[3]] = [
            [openplant_tbl[0], openplant_tbl[1], openplant_tbl[2]]
        ]
        matl_sql[openplant_tbl[3]] = []
    else:
        matl[openplant_tbl[3]].append(
            [openplant_tbl[0], openplant_tbl[1], openplant_tbl[2]]
        )


# Generate SQL statement for each table category

GENERAL_FIELD = (
    r"PlantArea, Guid, LineNumber, Specification, Rating, Schedule, Description"
)


def field_stm():
    # return r"PlantArea, Guid, LineNumber, Specification, Rating, Schedule, Description, InsulationMaterial, InsulationThickness"
    return r"PlantArea, Guid, LineNumber, Specification, Rating, Schedule, Description"


def size_stm(code):

    # ID	SIZE1	SIZE2	DESCRIPTION
    # 0	NominalDiameter	Nil	No 2 sizes, Size2 is not required
    # 1	NominalDiameter	NominalDiameterRunEnd	Size-2 is Run Side
    # 2	NominalDiameter	NominalDiameterBranchEnd	Size-2 is Branch Side (Tee)
    # 3	NominalDiameter	NominalDiameterBranch	Size-2 is Branch Side (O-Let)
    # 4	BoltDiameter	BoltLength	Bolt

    match int(code):
        case 0:
            sz1 = r"NominalDiameter"
            sz2 = r"Null"
        case 1:
            sz1 = r"NominalDiameter"
            sz2 = r"NominalDiameterRunEnd"
        case 2:
            sz1 = r"NominalDiameter"
            sz2 = r"NominalDiameterBranchEnd"
        case 3:
            sz1 = r"NominalDiameter"
            sz2 = r"NominalDiameterBranch"
        case 4:
            sz1 = r"BoltDiameter"
            sz2 = r"BoltLength"

    return f"{sz1} as Size1, {sz2} as Size2"


def qty_stm(code):

    match int(code):
        case 1:
            q = r"LengthEffective"
        case 5:
            q = r"NumberOfBolts"
        case _:
            q = 1
    return f"{q} as Qty"


for cat in matl_sql:
    if len(matl[cat]) > 0:
        sql_list = ""
        if len(matl[cat]) == 1:
            sql_list = f"SELECT '{matl[cat][0][0]}' as Ref_DB, {GENERAL_FIELD}, {size_stm(matl[cat][0][2])}, {qty_stm(matl[cat][0][1])} FROM {matl[cat][0][0]}"
        else:
            for idx in range(len(matl[cat])):
                sql_list += f"SELECT '{matl[cat][idx][0]}' as Ref_DB, {GENERAL_FIELD}, {size_stm(matl[cat][idx][2])}, {qty_stm(matl[cat][idx][1])} FROM {matl[cat][idx][0]}"
                if idx != len(matl[cat]) - 1:
                    sql_list += " UNION "

        # sql_sum = f"SELECT A.PlantArea, A.Description, A.Size1, A.Size2, sum(A.Qty) as NET FROM ({sql_list}) as A GROUP BY A.PlantArea, A.Description, A.Size1, A.Size2"

        matl_sql[cat] = sql_list



pipe_df_list = pd.read_sql(matl_sql["Pipe"], con=conn)
flange_df_list = pd.read_sql(matl_sql["Flange"], con=conn)
fitting_df_list = pd.read_sql(matl_sql["Fitting"], con=conn)
gasket_df_list = pd.read_sql(matl_sql["Gasket"], con=conn)
bolt_df_list = pd.read_sql(matl_sql["Bolt"], con=conn)
valve_df_list = pd.read_sql(matl_sql["Valve"], con=conn)
instr_df_list = pd.read_sql(matl_sql["Instr"], con=conn)
pipacc_df_list = pd.read_sql(matl_sql["PipAcc"], con=conn)

with pd.ExcelWriter("BQ.xlsx") as writer:
    pipe_df_list.to_excel(writer, sheet_name="pipe_list", index=False)
    flange_df_list.to_excel(writer, sheet_name="flg_list", index=False)
    fitting_df_list.to_excel(writer, sheet_name="fitting_list", index=False)
    gasket_df_list.to_excel(writer, sheet_name="gasket_list", index=False)
    bolt_df_list.to_excel(writer, sheet_name="bolt_list", index=False)
    valve_df_list.to_excel(writer, sheet_name="valve_list", index=False)
    instr_df_list.to_excel(writer, sheet_name="instr_list", index=False)
    pipacc_df_list.to_excel(writer, sheet_name="pipacc_list", index=False)



# Close connection
conn.close()
