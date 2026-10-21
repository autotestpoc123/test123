| 问题 | 当前代码与影响 | 修复方向 |
|---|---|---|
| **3. Blob 生命周期** | [user-photo.service.ts:101]（user-photo.service.ts:101) 淘汰缓存后仍固定延迟 5 秒 revoke，未判断是否仍被使用。在途条目被淘汰后才创建的 URL，也不受 `releaseAll()` 管理。 | 跟踪 URL 所有权；缓存淘汰且没有消费者使用时再释放，同时处理迟到响应。 |
| **4. 切换用户显示旧照片** | [user-view.component.ts:103](user-view.component.ts:103) 的 `populateUser()` 未清理照片订阅或关闭弹窗；照片回调也没有校验当前 MSID。 | 切换用户时取消旧订阅、关闭弹窗，并检查返回结果是否属于当前用户。 |
| **5. 旧定时器误删新缓存** | [user-photo.service.ts:80](/user-photo.service.ts:80) 仍使用 `setTimeout(() => this.cache.delete(key), …)`，可能删除同 key 后来生成的成功条目。 | 删除前检查 `this.cache.get(key) === entry`，只删除原条目。 |

第 5 项的最小修复是：
setTimeout(() => {
  if (this.cache.get(key) === entry) {
    this.cache.delete(key);
  }
}, NEGATIVE_CACHE_MS);
