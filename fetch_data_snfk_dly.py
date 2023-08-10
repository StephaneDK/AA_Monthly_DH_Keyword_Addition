# -*- coding: utf-8 -*-

from __future__ import print_function
import os
import sys


path_python = os.environ["path_code"].replace("Code","csv")
os.chdir(path_python)
#os.chdir('C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Auxilary tasks\\Amazon Advertisment\\Keyword Addition\\Monthly\\csv')

class LatestDataCheck(Exception):
    pass

#Getting last saturday date
from datetime import date
from datetime import timedelta
today = date.today()
last_saturday = today - timedelta(days= (today.weekday() - 5) % 7)
last_saturday = last_saturday.strftime('%Y-%m-%d')
#last_saturday = '2022-02-19' 


#Connecting to DB --------------------------------------------------------------------------------------------------------------------
#!/usr/bin/python
import snowflake.connector as sfc
import psycopg2
from config import Config
from configparser import ConfigParser
import csv



def config(filename='database.ini', section='snowflake_db'):
    # create a parser
    parser = ConfigParser()
    # read config file
    parser.read(filename)

    # get section, default to snowflake
    db = {}
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
    else:
        raise Exception('Section {0} not found in the {1} file'.format(section, filename))

    return db


def connect():
    """ Connect to the GDH server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the GDH server
        print('\nConnecting to the GDH database...')
        conn = sfc.connect(account='PRH',
                                region='us-east-1',
                                user= params['user'],
                                password= params['password'])
		
        # create a cursor
        cur = conn.cursor()
        
        #execute a statement
        print('Snowflake database version:')
        cur.execute('SELECT CURRENT_VERSION()')

        print('\nRetrieving data for forecasting')

        # display the Snowflake database server version
        db_version = cur.fetchone()
        print(db_version, ' \n ')
        

        
        sql_context = """
            with aa as (            
  
            select
                    
                    region_code, company, 
                    lower(keyword_text) as keyword,
                    sum(impressions) as impressions
                    
                    from PRH_GLOBAL.PUBLIC.X_REGION_AMS_SP_KEYWORDS_QUERY a
                    where company = 'DK'  
                    and to_varchar(report_date,'YYYY-MM') = to_varchar(current_date -30 ,'YYYY-MM')

                    group by 1,2,3
                    
            UNION ALL
            
            select                       
                    
                    region_code, company, 
                    lower(aat.query) as keyword,
                    sum(impressions) as impressions
                    
                    from PRH_GLOBAL.PUBLIC.X_REGION_AMS_SP_TARGETS aat
                    where company = 'DK' 
                    and to_varchar(report_date,'YYYY-MM') = to_varchar(current_date -30 ,'YYYY-MM')

                    group by 1,2,3
                    
            order by impressions desc
            
            ) 
            
            , ranks as (

                select 
                        SEARCH_TERM as keyword,
                        avg(SEARCH_FREQ_RANK) as SEARCH_FREQUENCY_RANK,
                        'US' as "REGION"
            
                from KEYWORD_WEBSEARCH.PUBLIC.AMZ_SEARCH_TERM_WK
            
                where category = 'All' 
                and to_varchar(WEEK_END_DATE,'YYYY-MM') = to_varchar(current_date -30 ,'YYYY-MM')
                group by 1,3
                

                UNION ALL

                select 
                        SEARCH_TERM as keyword, 
                        avg(SEARCH_FREQUENCY_RANK) as SEARCH_FREQUENCY_RANK, 
                        'UK' as "REGION"
                        
                from PRH_GLOBAL_UK_SANDBOX..GDH_EDW_AMAZON_SEARCH_TERMS 
            
                where   
                    region = 'GB'
                    and summary_level = 'Weekly' 
                    and department = 'Amazon.co.uk'
                    and to_varchar(TO_DATE(SRC_FILE_DATE),'YYYY-MM') = to_varchar(current_date -30 ,'YYYY-MM')

                group by 1,3
                
                order by 2 desc

                    
            )

            ,DH_keywords as (

                select distinct keyword, 'Y' as "Datahawk" from DATAHAWK_SHARE_V2.REFERENTIAL.REFERENTIAL_KEYWORD_TRACKED

            )

                
                select
                        b."REGION", 
                        a.company, 
                        b.keyword,
                        b.SEARCH_FREQUENCY_RANK,
                        c.volume,
                        a.impressions,  --sum(a.impressions),
                        e."Datahawk"
                            

                from ranks b

                left join DH_keywords e on b.keyword = e.keyword
                
                left join aa a on b.keyword = a.keyword and b."REGION" = a.region_code

                left join PRH_GLOBAL_DK_SANDBOX.PUBLIC.AMZ_SEARCH_RANK_VOLUME c on round(b.SEARCH_FREQUENCY_RANK) = c.rank and lower(b."REGION") = c.country  

                where  
                    b.keyword IS NOT NULL
                    --and e."Datahawk" IS NOT NULL
                    --and a.region_code = 'US' 
                    

                --group by 1,2,3,4,5,7
                order by 4 nulls last

        """
        
        cur.execute(sql_context)
    
        # Fetch all rows from database
        record = cur.fetchall()
            
        #Writing csv file
        with open('New Terms DH.csv', 'w', encoding="utf-8" ) as f:
            column_names = [i[0] for i in cur.description]
            file = csv.writer(f, lineterminator = '\n')
            file.writerow(column_names)
            file.writerows(record)
            print('Latest Search Ranks saved \n')



        #close the communication with the PostgreSQL
        cur.close()
        

        
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Keyword data retrieved succesfully\n')

if __name__ == '__main__':
    connect()








