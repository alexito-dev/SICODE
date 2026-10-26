import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import {
  FaCircleCheck, FaCircleXmark, FaCloudArrowUp, FaDownload, FaFileCode, FaFileExcel,
  FaFileImage, FaFileLines, FaFilePdf, FaFileWord, FaFileZipper, FaLink, FaMagnifyingGlass,
  FaTrashCan, FaUser,
} from 'react-icons/fa6'
import { deleteFile, listFiles, uploadFiles } from './api.js'

const ICONS = {
  zip: [FaFileZipper, ['zip', 'rar', '7z', 'tar', 'gz', 'iso']],
  pdf: [FaFilePdf, ['pdf']],
  excel: [FaFileExcel, ['xls', 'xlsx', 'xlsm', 'csv', 'ods']],
  word: [FaFileWord, ['doc', 'docx', 'odt', 'rtf', 'txt']],
  image: [FaFileImage, ['png', 'jpg', 'jpeg', 'gif', 'svg', 'webp', 'bmp']],
  code: [FaFileCode, ['exe', 'msi', 'bat', 'py', 'js', 'json', 'sql', 'xml']],
}

function iconOf(name) {
  const ext = name.includes('.') ? name.split('.').pop().toLowerCase() : ''
  const kind = Object.keys(ICONS).find((k) => ICONS[k][1].includes(ext))
  return { Icon: kind ? ICONS[kind][0] : FaFileLines, kind: kind || 'other' }
}

const UNITS = ['B', 'KB', 'MB', 'GB', 'TB']

function formatSize(n) {
  let i = 0
  while (n >= 1024 && i < 4) { n /= 1024; i++ }
  return `${i ? n.toFixed(1) : n} ${UNITS[i]}`
}

function copyText(text) {
  if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text)
  const t = document.createElement('textarea')
  t.value = text
  document.body.append(t)
  t.select()
  document.execCommand('copy')
  t.remove()
  return Promise.resolve()
}

function FileRow({ file, onCopy, onDelete }) {
  const { Icon, kind } = iconOf(file.name)
  return (
    <div className="row">
      <div className={`ficon ${kind}`}><Icon /></div>
      <div className="info">
        <div className="name" title={file.name}>{file.name}</div>
        <div className="meta">
          {formatSize(file.size)} · {new Date(file.date * 1000).toLocaleDateString('es-PE')}
          {file.by ? ` · ${file.by}` : ''}
        </div>
      </div>
      <button className="ib" onClick={() => onCopy(file)} title="Copiar enlace"><FaLink /></button>
      <a className="ib dl" href={`/d/${file.id}`} title="Descargar"><FaDownload /></a>
      <button className="ib del" onClick={() => onDelete(file)} title="Eliminar"><FaTrashCan /></button>
    </div>
  )
}

export default function App() {
  const [files, setFiles] = useState([])
  const [query, setQuery] = useState('')
  const [progress, setProgress] = useState(null)
  const [over, setOver] = useState(false)
  const [toast, setToast] = useState(null)
  const [user, setUser] = useState('')
  const input = useRef(null)

  const notify = (msg, error = false) => {
    setToast({ msg, error })
    setTimeout(() => setToast(null), 2500)
  }

  const load = useCallback(async () => {
    try {
      setFiles(await listFiles())
    } catch (e) {
      notify(e.message, true)
    }
  }, [])

  useEffect(() => {
    fetch('/api/me').then((r) => (r.ok ? r.json() : null)).then((u) => u && setUser(u.nombre)).catch(() => {})
  }, [])

  useEffect(() => {
    load()
    const t = setInterval(load, 15000)
    return () => clearInterval(t)
  }, [load])

  const upload = async (list) => {
    if (!list.length) return
    setProgress(0)
    try {
      await uploadFiles(list, setProgress)
      notify('Subido')
    } catch (e) {
      notify(e.message, true)
    }
    setProgress(null)
    load()
  }

  const copy = async (file) => {
    await copyText(`${location.origin}/d/${file.id}`)
    notify('Enlace copiado')
  }

  const remove = async (file) => {
    if (!confirm(`Eliminar "${file.name}"?`)) return
    try {
      await deleteFile(file.id)
    } catch (e) {
      notify(e.message, true)
    }
    load()
  }

  const rows = useMemo(() => {
    const q = query.toLowerCase()
    return files.filter((f) => f.name.toLowerCase().includes(q))
  }, [files, query])

  return (
    <>
      <header>
        <img src="/logo.png" alt="" />
        <span>RepoSUNAT</span>
        {user && <small><FaUser /> {user}</small>}
      </header>

      <main>
        <div
          className={`drop${over ? ' over' : ''}`}
          onClick={() => input.current.click()}
          onDragOver={(e) => { e.preventDefault(); setOver(true) }}
          onDragLeave={() => setOver(false)}
          onDrop={(e) => { e.preventDefault(); setOver(false); upload([...e.dataTransfer.files]) }}
        >
          <FaCloudArrowUp />
          <div>Suelta tus archivos aqui</div>
          {progress !== null && <div className="bar"><i style={{ width: `${progress * 100}%` }} /></div>}
        </div>
        <input
          ref={input}
          type="file"
          multiple
          hidden
          onChange={(e) => { upload([...e.target.files]); e.target.value = '' }}
        />

        {files.length > 0 && (
          <div className="search">
            <FaMagnifyingGlass />
            <input
              type="text"
              placeholder="Buscar"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
        )}

        <div className="list">
          {rows.map((f) => <FileRow key={f.id} file={f} onCopy={copy} onDelete={remove} />)}
        </div>
      </main>

      <div className={`toast${toast ? ' show' : ''}${toast?.error ? ' err' : ''}`}>
        {toast?.error ? <FaCircleXmark /> : <FaCircleCheck />}
        <span>{toast?.msg}</span>
      </div>
    </>
  )
}
