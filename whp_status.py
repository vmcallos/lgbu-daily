import os
from datetime import datetime
from datetime import timedelta
import pandas as pd
import re

def main():
    # Get filenames of Daily PowerPlant Report
    dir_path = os.getcwd()
    path = os.path.join(dir_path, 'data')
    for root,dirs,files in os.walk(dir_path):
        xlsfiles=[ _ for _ in files if _.endswith('.xlsx') |  _.endswith('.xls')]

    for idx, xls in enumerate(xlsfiles):
        print(idx, xls)
        # Create shift A,B and C from filename
        shifts = create_shift(xls)

        if idx==0:
            # Read Malitbog sheet
            malitbog_df = pd.read_excel(path+'\\'+xls, sheet_name='malitbog', usecols=[0, 1, 2, 6, 7, 11, 12],
                           skiprows=4, nrows=24, index_col=0)
            malitbog_status, malitbog_whp = whp_status(malitbog_df, shifts)

            # Read South Sambaloran sheet
            ssamb_df = pd.read_excel(path+'\\'+xls, sheet_name='so.samb.', usecols=[0, 1, 2, 6, 7, 11, 12],
                           skiprows=3, nrows=23, index_col=0)
            ssamb_status, ssamb_whp = whp_status(ssamb_df, shifts)

        else:
            malitbog_df = pd.read_excel(path+'\\'+xls, sheet_name='malitbog', usecols=[0, 1, 2, 6, 7, 11, 12],
                           skiprows=4, nrows=24, index_col=0)
            temp_status, temp_whp = whp_status(malitbog_df, shifts)
            # Concatenate whp and status data
            malitbog_status = pd.concat([malitbog_status, temp_status])
            malitbog_whp = pd.concat([malitbog_whp, temp_whp])

            # Read South Sambaloran sheet
            ssamb_df = pd.read_excel(path+'\\'+xls, sheet_name='so.samb.', usecols=[0, 1, 2, 6, 7, 11, 12],
                           skiprows=3, nrows=23, index_col=0)
            temp_status, temp_whp = whp_status(ssamb_df, shifts)
            # Concatenate whp and status data to main DataFrame
            ssamb_status = pd.concat([ssamb_status, temp_status])
            ssamb_whp = pd.concat([ssamb_whp, temp_whp])

    # Export DataFrames to csv
    malitbog_whp.to_csv('malitbog_whp.csv')
    malitbog_status.to_csv('malitbog_status.csv')
    ssamb_whp.to_csv('ssamb_whp.csv')
    ssamb_status.to_csv('ssamb_status.csv')

#####################################################################################################

def create_shift(xls):
    #Get date from filename
    match = re.search(r'\d+?\d*', xls)
    date = match.group()
    today = datetime.strptime(date, '%m%d%y')

    # Create Series for 3-shifts >>will be used as index
    shift = pd.Series([today + timedelta(hours=h) for h in [0,8,16] ]) # veryfy if 0,8,16 or 8,16,24

    return shift

def whp_status(df, shift_abc):
    # Transpose the DataFrame; wellname as column
    dft = df.transpose()

    # Select the status rows for shifts A,B,C
    status = dft.iloc[[0,2,4], :]
    # Set index for shift A,B,C timestamp
    status = status.set_index(shift_abc)

    # Select the whp rows for shifts A,B,C
    whp = dft.iloc[[1,3,5], :]
    # Set index for shift A,B,C timestamp
    whp = whp.set_index(shift_abc)

    return(status, whp)


if __name__ == '__main__':
    main()
