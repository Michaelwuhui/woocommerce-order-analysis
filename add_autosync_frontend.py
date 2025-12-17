
import os
import re

file_path = r'c:\Users\Administrator\Documents\GitHub\woocommerce-order-analysis\templates\settings.html'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add Toggle Switch HTML
# We'll add it in the header, next to the buttons.
# Current header:
# <div class="d-flex gap-2">
#     <button ... id="syncAllBtn">...</button>
#     <button ...>...</button>
# </div>

toggle_html = """        <div class="d-flex align-items-center me-3">
            <div class="form-check form-switch">
                <input class="form-check-input" type="checkbox" id="autoSyncToggle">
                <label class="form-check-label text-white-50" for="autoSyncToggle">自动同步 (每15分钟)</label>
            </div>
        </div>"""

# We want to insert this before the buttons container.
# Let's look for <div class="d-flex gap-2">
if '<div class="d-flex gap-2">' in content:
    content = content.replace('<div class="d-flex gap-2">', toggle_html + '\n        <div class="d-flex gap-2">')

# 2. Add JavaScript Logic
# We need to fetch status on load and handle toggle change.
js_logic = """
    // Auto Sync Toggle
    const autoSyncToggle = document.getElementById('autoSyncToggle');
    if (autoSyncToggle) {
        // Fetch initial status
        fetch('/api/settings/autosync')
        .then(res => res.json())
        .then(data => {
            autoSyncToggle.checked = data.enabled;
        });
        
        // Handle change
        autoSyncToggle.addEventListener('change', function() {
            const enabled = this.checked;
            fetch('/api/settings/autosync', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ enabled: enabled })
            })
            .then(res => res.json())
            .then(data => {
                if (data.success) {
                    console.log('Auto sync set to:', data.enabled);
                } else {
                    alert('设置失败');
                    this.checked = !enabled; // Revert
                }
            })
            .catch(err => {
                console.error('Error setting auto sync:', err);
                alert('设置错误');
                this.checked = !enabled; // Revert
            });
        });
    }
"""

# Insert JS logic inside DOMContentLoaded
# We can insert it after "Sync All Button" block
if "// Sync All Button" in content:
    content = content.replace("// Sync All Button", js_logic + "\n    // Sync All Button")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Successfully added auto-sync toggle to settings.html")
