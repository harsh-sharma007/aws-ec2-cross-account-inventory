# AWS Multi-Account EC2 Inventory Automation

![AWS](https://img.shields.io/badge/AWS-%23FF9900.svg?style=for-the-badge&logo=amazon-aws&logoColor=white)
![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)

An automated, serverless Python solution that gathers EC2 inventory across multiple AWS accounts and regions. It extracts resource metadata, computes 30-day historical CPU utilization metrics, and dispatches a dynamically formatted, multi-sheet Excel report via Amazon SES. 

Additionally, it features a passive auditing mechanism to detect untagged instances launched by Auto Scaling Groups (ASGs).

---

## System Architecture

This solution uses an event-driven, cross-account fan-out model executed entirely in memory:

1. **Host Execution (Hub):** An Amazon EventBridge schedule triggers the AWS Lambda function in the central host account.
2. **Cross-Account Handshake (Spokes):** The Lambda function assumes IAM roles (`sts:AssumeRole`) in predefined target accounts using Boto3.
3. **Data Aggregation:** The function queries `ec2:DescribeInstances` and `cloudwatch:GetMetricStatistics` across all active regions in the target accounts.
4. **Data Formatting:** The raw data is converted into a Pandas DataFrame and formatted into a styled `.xlsx` workbook using `openpyxl` (entirely in an `io.BytesIO()` memory stream).
5. **Email Dispatch:** The generated workbook is attached to an HTML-formatted email and dispatched to stakeholders via Amazon SES.

---

## Prerequisites & Requirements

### 1. AWS Account Setup
* **Host Account (Hub):** The central account hosting the Lambda function and SES identity.
* **Target Accounts (Spokes):** The secondary accounts containing the EC2 resources to be scanned.

### 2. Amazon SES Setup
* The sender email address MUST be a verified identity in Amazon SES in the Host Account.
* If sending to external recipients, ensure your AWS SES account is moved out of the Sandbox environment.

### 3. Lambda Configuration
Due to the computational overhead of processing thousands of instances and building Excel files, configure the Lambda function with the following specifications:
* **Runtime:** Python 3.9 / 3.10 / 3.11
* **Timeout:** `15 minutes` (900 seconds) - *Required for synchronous CloudWatch API calls.*
* **Memory (RAM):** `1024 MB` to `2048 MB`

---

## Environment Variables

The Lambda function requires the following environment variables:

| Variable | Description | Example |
| :--- | :--- | :--- |
| `SENDER_EMAIL` | Verified SES sender email address | `reports@yourdomain.com` |
| `RECIPIENT_EMAILS`| Comma-separated list of recipient emails | `team@yourdomain.com, audit@yourdomain.com` |
| `TARGET_ACCOUNTS` | Comma-separated list of 12-digit AWS Account IDs | `111122223333, 444455556666` |
| `SES_REGION` | AWS Region where SES identity is verified | `ap-south-1` |
| `ROLE_NAME` | Name of cross-account IAM role in target accounts | `EC2InventoryRole` |

---

## Deployment & Dependencies

### Python Libraries (Lambda Layer)
Because standard AWS Python runtimes do not include data science or spreadsheet libraries, you must deploy this function with the following dependencies (either bundled in a `.zip` or via an AWS Lambda Layer):
* `pandas`
* `openpyxl`
* *Note: `boto3` is included natively in the AWS Lambda runtime and does not need to be bundled.*

---

## IAM Security Setup

### Host Account Permissions (Lambda Execution Role)
Attach the `AWSLambdaBasicExecutionRole` managed policy (for CloudWatch Logging), plus the following custom inline policy:

```json
{
    "Version": "2012-10-17",
    "Statement": [
        {
            "Sid": "AllowCrossAccountRoleAssumption",
            "Effect": "Allow",
            "Action": "sts:AssumeRole",
            "Resource": "arn:aws:iam::*:role/EC2InventoryRole"
        },
        {
            "Sid": "AllowSESEmailDispatch",
            "Effect": "Allow",
            "Action": "ses:SendRawEmail",
            "Resource": "*"
        }
    ]
}


### Prompts used


"Develop an automated solution that gathers EC2 inventory across multiple AWS accounts and regions using AWS Organizations or AssumeRole. Collect instance details, OS, IPS, AMIS, security groups, EBS volumes, backup status, ASG membership, tags, and resource utilization. Automatically identify Auto Scaling instances without Name tags and populate names using launch template or ASG tags. Consolidate data into a standardized Excel workbook with multiple worksheets and automatically email the final report to clients every month."

"i don't want to use ecs fargate"

"i want to use lambda only, and i want one workbook in which there will be sheets per account in which the inventory will be there"

"what iam permissions should i give to the lambda function"

"is there any other way to create this automation"

"how will the report look like if we use the lambda function"

"this will work with multiple accounts in the organization right?"

"will it work if we have like 9 accounts in the org and like 200 servers per account"

"it doesn't provide the ebs volumes attached, asg membership, tags per instance?"

"how will the sheet look like now?"

"what will be the associated cost?"

"outside of free tier?"

"can i also create the lambda function that i can pass the account id for which i want the lambda to fetch the ec2 inventory?"

"but i want this to work for all the accounts when it runs automatically every month"

"[WARNING] 2026-07-15T18:09:52.343Z LAMBDA_WARNING... [ERROR] Runtime.ImportModuleError: Unable to import module 'lambda_function': No module named 'pandas'" (Pasted Error Log)


"[WARNING] 2026-07-15T18:12:47.849Z LAMBDA_WARNING... [ERROR] AttributeError: 'None Type' object has no attribute 'split'" (Pasted Error Log)


"[WARNING] 2026-07-15T18:14:17.453Z LAMBDA_WARNING... [ERROR] AttributeError: 'NoneType' object has no attribute 'split'" (Pasted Error Log)


"currently i am testing in a single aws account, there is no aws org setup, so while testing the lambda function is failing"

"currently i want it to work like the accounts in which i configure the assume role for the lambda function, for those accounts only i should get the inventory"

"will it also work if i don't specify the accounts and it can automaticall fetch the data of the accounts on which the assume role has been created?"

"lets suppose i want the inventroy of the account in which i am running the lambda function then will it work for the account?"

"give the code"

"the lambda function executed but i didn't receive any excel sheet"

"i am still not getting any email attachment"

"give me the updated code and the region we are using is ap-south-1"

"it worked now but didn't include the security group, tags"

"IT STILL DIDN'T HAVE THE SECURITY GROUP DETAILS"

"okay perfect now i want you to formatize the email a lil bit"

"make the subject AWS EC2 Inventory"

"okay now i have another account for testing so give me the assume role configuration now"

"Add this block to your Lambda Execution Role Policy: JSON { "Sid": 
"AllowCrossAccountRoleAssumption", "Effect": "Allow", "Action": "sts:AssumeRole", "Resource": "arn:aws:iam::*:role/EC2InventoryRole" } WHERE TO ADD THIS?"

"GIVE A SEPARATE POLICY"

"okay its working fine now for multi account too, now i want to make the excel sheet look a lil better, in every sheet, there should be the AWS Account name mentioned above the content"

"okay now a few changes that i need are borders to be added the inventory subject should have the name of the month for which the inventory report is being generated"

"the excel report is like that, i have to manually stretch the columns to see the data"

"the data should be centre aligned"

"okay all sorted till now moving on the next step we want to identify the instances which don't have any name tag and are launched by an autoscaling group and list those resources after the inventory list for the ec2 instances, i want the details like instance id, name, private ip of those instances"

"it didn't show any data in the sheet for the untagged instances which are a part of an asg"

"don't add the asg name in the name column where there is no name just leave it blank, we will add the name manually from our side in the sheet also add the status, region, instance type for these instances"

"this time the excel sheet didn't had any table for the instances without any name tags"

"i don't want to create any tags, i just want the report for the instances without the name tags along with the inventory"

"also include the asg name in the instances with no tag table"

"what libraries are we using here and what is their function"

"can you explain how the iam policy for cross account is working here"

"it takes like 30 seconds for this lambda function to run, will it run when i have like 10 aws accounts and 200 servers/account"

