import os
from datetime import datetime
import sys

def main(sonar_cloud_project_name):
    
    readme = "./README.md"
    
    current_date = datetime.now()
    
    sonar_cloud_url = "https://sonarcloud.io/dashboard?id="
    sonar_report_url = sonar_cloud_url + sonar_cloud_project_name
    
    data = "* Last Repository check was made at: {} \n Check code analysis results at:{}".format(current_date, sonar_report_url)

    with open(readme, "a") as f:
        print(data, file=f)

if __name__ == "__main__":
    main(sys.argv[0])
    
