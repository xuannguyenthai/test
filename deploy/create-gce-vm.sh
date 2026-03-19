#!/usr/bin/env bash
#
# Creates a GCE Windows Server VM for lana-excel.
#
# Prerequisites:
#   - gcloud CLI authenticated (gcloud auth login)
#   - A GCP project selected (gcloud config set project YOUR_PROJECT)
#
# Usage:
#   chmod +x deploy/create-gce-vm.sh
#   ./deploy/create-gce-vm.sh
#
# After the VM is created:
#   1. Set the Windows password:  gcloud compute reset-windows-password lana-excel-vm
#   2. RDP into the VM
#   3. Clone/copy the repo to C:\lana-excel
#   4. Run: .\deploy\install-office.ps1   (installs Excel)
#   5. Run: .\deploy\setup-gce-vm.ps1     (installs Python, deps, service)
#   6. Start-Service lana-excel-watcher

set -euo pipefail

# ---------------------------------------------------------------------------
# Configuration — override via environment variables
# ---------------------------------------------------------------------------
VM_NAME="${LANA_VM_NAME:-lana-excel-vm}"
ZONE="${LANA_ZONE:-asia-southeast1-b}"
MACHINE_TYPE="${LANA_MACHINE_TYPE:-n2-standard-4}"    # 4 vCPU, 16 GB RAM
BOOT_DISK_SIZE="${LANA_DISK_SIZE:-100GB}"
IMAGE_FAMILY="windows-2025"
IMAGE_PROJECT="windows-cloud"
NETWORK_TAG="lana-excel"

echo "=== Creating GCE Windows VM for lana-excel ==="
echo ""
echo "  VM name:      $VM_NAME"
echo "  Zone:         $ZONE"
echo "  Machine type: $MACHINE_TYPE"
echo "  Disk size:    $BOOT_DISK_SIZE"
echo "  Image:        $IMAGE_FAMILY ($IMAGE_PROJECT)"
echo ""

# ---------------------------------------------------------------------------
# Create firewall rule for RDP (if not exists)
# ---------------------------------------------------------------------------
gcloud services enable compute.googleapis.com

if ! gcloud compute firewall-rules describe allow-rdp --quiet 2>/dev/null; then
    echo "[1/3] Creating firewall rule for RDP ..."
    gcloud compute firewall-rules create allow-rdp \
        --direction=INGRESS \
        --priority=1000 \
        --network=default \
        --action=ALLOW \
        --rules=tcp:3389 \
        --source-ranges=0.0.0.0/0 \
        --target-tags="$NETWORK_TAG" \
        --description="Allow RDP access to lana-excel VMs"
else
    echo "[1/3] Firewall rule 'allow-rdp' already exists."
fi

# ---------------------------------------------------------------------------
# Create the VM
# ---------------------------------------------------------------------------
echo ""
echo "[2/3] Creating VM: $VM_NAME ..."

gcloud compute instances create "$VM_NAME" \
    --zone="$ZONE" \
    --machine-type="$MACHINE_TYPE" \
    --image-family="$IMAGE_FAMILY" \
    --image-project="$IMAGE_PROJECT" \
    --boot-disk-size="$BOOT_DISK_SIZE" \
    --boot-disk-type=pd-ssd \
    --tags="$NETWORK_TAG" \
    --scopes=default,storage-ro \
    --metadata sysprep-specialize-script-cmd="googet -noconfirm=true install google-compute-engine-ssh",enable-windows-ssh=TRUE

echo ""
echo "[3/3] VM created successfully."

# ---------------------------------------------------------------------------
# Get the external IP
# ---------------------------------------------------------------------------
EXTERNAL_IP=$(gcloud compute instances describe "$VM_NAME" \
    --zone="$ZONE" \
    --format="get(networkInterfaces[0].accessConfigs[0].natIP)")

echo ""
echo "=== VM Ready ==="
echo ""
echo "  External IP: $EXTERNAL_IP"
echo ""
echo "  Next steps:"
echo "    1. Set Windows password:"
echo "         gcloud compute reset-windows-password $VM_NAME --zone=$ZONE"
echo ""
echo "    2. RDP into the VM:"
echo "         Use Remote Desktop to connect to $EXTERNAL_IP"
echo ""
echo "    3. Copy lana-excel to the VM:"
echo "         gcloud compute scp --recurse --zone=$ZONE ./  ${VM_NAME}:C:\\\\lana-excel\\\\"
echo ""
echo "    4. Open PowerShell as Administrator on the VM and run:"
echo "         cd C:\\lana-excel"
echo "         .\\deploy\\install-office.ps1"
echo "         .\\deploy\\setup-gce-vm.ps1"
echo ""
echo "    5. Start the service:"
echo "         Start-Service lana-excel-watcher"
echo ""
echo "  Estimated monthly cost: ~\$150-200 (n2-standard-4 + Windows license)"
echo "  To save costs, stop the VM when not in use:"
echo "    gcloud compute instances stop $VM_NAME --zone=$ZONE"
