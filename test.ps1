Test Case ID,Feature,Description,Steps,Expected Result,Actual Result,Priority
NOTE-001,Note,Create a note in external group,1. Open external group chat
2. Tap "Note" or "Create Note"
3. Add title, content, images/files
4. Save note,Note is saved and visible in external group; all members can view it,,High
NOTE-002,Note,Forward note from external to internal group,1. Open note in external group
2. Tap "Forward/Share"
3. Select internal group
4. Confirm send,Note appears in internal group with correct content, attachments, and timestamp,,High
NOTE-003,Note,Edit note in external group after forwarding,1. Edit note in external group
2. Observe internal group,Behavior matches spec: either synced or not; no data corruption,,Medium
NOTE-004,Note,Comment on forwarded note in internal group,1. Open forwarded note in internal group
2. Add comment,Comment visible to internal members; external group remains unchanged,,Medium
NOTE-005,Note,Forward note with large attachments,1. Create note with large files (>50MB)
2. Forward to internal group,Upload succeeds or error message displayed,,High
NOTE-006,Note,Forward note without permission,1. User without internal group permission attempts forwarding,Forwarding blocked with notification,,High
NOTE-007,Note,Cross-platform forward,1. Forward note from mobile external group
2. Open internal group on desktop,Content consistent across platforms,,Medium
NOTE-008,Note,Create note with invalid content,1. Attempt note with empty content, invalid characters, or unsupported file type,Note creation fails; error message shown,,Medium
SOL-001,Solitaire,Launch Solitaire in external group,1. Open external group
2. Tap "Solitaire"
3. Start new game,Game opens and is playable,,High
SOL-002,Solitaire,Forward Solitaire game to internal group,1. Tap "Forward/Share" on active game
2. Select internal group,Internal group receives game invite; members can join,,High
SOL-003,Solitaire,Forward ongoing game mid-progress,1. Play game in external group
2. Forward to internal group,Internal group receives latest game state; progress consistent,,High
SOL-004,Solitaire,Multiple game instances forwarded,1. Multiple external groups forward same game to internal group,Internal group handles instances separately; no data clash,,Medium
SOL-005,Solitaire,Game notifications,1. Forward game to internal group
2. Observe notifications,Internal group members receive correct notifications; no duplicates,,Medium
SOL-006,Solitaire,Permission restriction on forwarded game,1. Non-member of internal group tries interaction,Interaction blocked; proper message displayed,,High
SOL-007,Solitaire,Interrupted forward,1. Start forward while network disconnects,Error message displayed; game state not corrupted,,Medium
SOL-008,Solitaire,Forward from mobile to desktop,1. Forward game from mobile external group
2. Access internal group on desktop,Game state consistent; UI renders correctly,,Medium
