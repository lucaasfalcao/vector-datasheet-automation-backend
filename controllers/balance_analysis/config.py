"""Tags description to Balance Analysis API"""

from fastapi import APIRouter

BALANCE_ANALYSIS = dict(
    name='Balance Analysis',
    description='API to manage balance analysis operations.'
)


tags = [BALANCE_ANALYSIS]

router = APIRouter()
