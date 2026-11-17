const DATABASE_NAME = 'gridnova-workbooks'
const DATABASE_VERSION = 1
const STORE_NAME = 'workbooks'

function openDatabase() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('This browser does not support IndexedDB.'))
      return
    }
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION)
    request.onupgradeneeded = () => {
      const database = request.result
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        const store = database.createObjectStore(STORE_NAME, { keyPath: 'id' })
        store.createIndex('updatedAt', 'updatedAt', { unique: false })
      }
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error || new Error('Could not open workbook storage.'))
    request.onblocked = () => reject(new Error('Workbook storage is blocked by another browser tab.'))
  })
}

function requestResult(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error || new Error('Workbook storage request failed.'))
  })
}

async function withStore(mode, callback) {
  const database = await openDatabase()
  try {
    const transaction = database.transaction(STORE_NAME, mode)
    const store = transaction.objectStore(STORE_NAME)
    const request = callback(store)
    const result = await requestResult(request)
    await new Promise((resolve, reject) => {
      transaction.oncomplete = resolve
      transaction.onerror = () => reject(transaction.error || new Error('Workbook storage transaction failed.'))
      transaction.onabort = () => reject(transaction.error || new Error('Workbook storage transaction was aborted.'))
    })
    return result
  } finally {
    database.close()
  }
}

export function saveWorkbook(record) {
  if (!record?.id || !record?.snapshot) {
    return Promise.reject(new Error('A workbook ID and snapshot are required to save.'))
  }
  return withStore('readwrite', (store) => store.put({ ...record, updatedAt: Date.now() }))
}

export function getWorkbook(id) {
  if (!id) return Promise.resolve(null)
  return withStore('readonly', (store) => store.get(id))
}

export function listWorkbooks() {
  return withStore('readonly', (store) => store.getAll()).then((records) =>
    (records || []).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0)),
  )
}

export function deleteWorkbook(id) {
  if (!id) return Promise.resolve()
  return withStore('readwrite', (store) => store.delete(id))
}
