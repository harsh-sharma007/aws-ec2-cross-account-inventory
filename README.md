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
