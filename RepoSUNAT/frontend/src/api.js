export async function listFiles() {
  const r = await fetch('/api/files')
  if (!r.ok) throw new Error('No se pudo cargar la lista')
  return r.json()
}

export async function deleteFile(id) {
  const r = await fetch(`/api/files/${id}`, { method: 'DELETE' })
  if (!r.ok) throw new Error('No se pudo borrar')
}

export function uploadFiles(files, onProgress) {
  return new Promise((resolve, reject) => {
    const fd = new FormData()
    for (const f of files) fd.append('files', f)
    const x = new XMLHttpRequest()
    x.upload.onprogress = (e) => e.lengthComputable && onProgress(e.loaded / e.total)
    x.onload = () => (x.status === 200 ? resolve() : reject(new Error('Error al subir')))
    x.onerror = () => reject(new Error('Error al subir'))
    x.open('POST', '/api/upload')
    x.send(fd)
  })
}
