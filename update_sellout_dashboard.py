"""
eXtra Sellout Dashboard - Auto Updater
=======================================
새 주간 데이터 감지 → data.json 재생성 → GitHub Pages 자동 배포.

Usage:
    python update_sellout_dashboard.py          # 새 주차 감지 후 업데이트
    python update_sellout_dashboard.py --force  # 강제 업데이트
"""

import os
import sys
import json
import glob
import subprocess
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WEEKLY_DIR = r"C:\Users\J_park\Documents\2026\01. Work\01. Sales\01. Sell out\01. Weekly\01. eXtra Raw\01. Sell out\00. Weekly"
DATA_JSON_PATH = os.path.join(SCRIPT_DIR, "data.json")
STATE_FILE = os.path.join(SCRIPT_DIR, ".update_state.json")


def get_available_weeks():
    """사용 가능한 주간 파일 목록"""
    weeks = []
    for f in sorted(glob.glob(os.path.join(WEEKLY_DIR, "week*.xlsx"))):
        fname = os.path.basename(f)
        try:
            num = int(fname.replace("week", "").replace(".xlsx", ""))
            weeks.append(num)
        except ValueError:
            pass
    return weeks


def get_last_processed_week():
    """마지막 처리된 주차"""
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f:
            state = json.load(f)
        return state.get("last_week", 0)
    return 0


def save_state(last_week):
    """처리 상태 저장"""
    state = {
        "last_week": last_week,
        "updated_at": datetime.now().isoformat(),
    }
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)


def generate_data():
    """data.json 재생성"""
    print("\n[STEP 1] Generating data.json...")
    result = subprocess.run(
        [sys.executable, os.path.join(SCRIPT_DIR, "generate_sellout_data.py")],
        cwd=SCRIPT_DIR,
        capture_output=True, text=True, timeout=600
    )
    print(result.stdout)
    if result.returncode != 0:
        print(f"ERROR: {result.stderr}")
        return False
    return True


def deploy_to_github():
    """GitHub Pages 배포"""
    print("\n[STEP 2] Deploying to GitHub Pages...")
    os.chdir(SCRIPT_DIR)

    try:
        # Check if git repo
        result = subprocess.run(["git", "status"], capture_output=True, text=True, cwd=SCRIPT_DIR)
        if result.returncode != 0:
            print("  Not a git repo, skipping deploy")
            return False

        # Add data.json
        subprocess.run(["git", "add", "data.json"], cwd=SCRIPT_DIR)

        # Check if there are changes
        result = subprocess.run(["git", "diff", "--cached", "--name-only"], capture_output=True, text=True, cwd=SCRIPT_DIR)
        if not result.stdout.strip():
            print("  No changes to deploy")
            return True

        # Commit
        week = max(get_available_weeks())
        msg = f"Update sellout data W{week:02d} ({datetime.now().strftime('%Y-%m-%d')})"
        subprocess.run(["git", "commit", "-m", msg], cwd=SCRIPT_DIR)

        # Push
        result = subprocess.run(["git", "push"], capture_output=True, text=True, cwd=SCRIPT_DIR)
        if result.returncode == 0:
            print(f"  Deployed: {msg}")
            return True
        else:
            print(f"  Push failed: {result.stderr}")
            return False

    except Exception as e:
        print(f"  Deploy error: {e}")
        return False


def main():
    force = "--force" in sys.argv

    print("=" * 60)
    print("eXtra Sellout Dashboard - Auto Updater")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # Check for new weeks
    available = get_available_weeks()
    last_processed = get_last_processed_week()

    if not available:
        print("No weekly files found.")
        return

    latest = max(available)
    print(f"Available weeks: W{min(available):02d}~W{latest:02d}")
    print(f"Last processed: W{last_processed:02d}")

    if latest <= last_processed and not force:
        print(f"\nNo new data (latest W{latest:02d} already processed)")
        return

    new_weeks = [w for w in available if w > last_processed]
    if new_weeks:
        print(f"New weeks to process: {', '.join(f'W{w:02d}' for w in new_weeks)}")
    elif force:
        print("Force update requested")

    # Generate
    if not generate_data():
        print("\nFAILED: Data generation failed")
        return

    # Deploy
    deployed = deploy_to_github()

    # Save state
    save_state(latest)

    # Summary
    size_mb = os.path.getsize(DATA_JSON_PATH) / 1024 / 1024
    print("\n" + "=" * 60)
    print(f"DONE: W{latest:02d} processed, data.json={size_mb:.1f}MB" +
          (", deployed" if deployed else ", deploy skipped"))
    print("=" * 60)


if __name__ == "__main__":
    main()
