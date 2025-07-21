import { ZERO } from '../../../dataset/constant/Common'
import { VIRTUAL_ELEMENT_TYPE } from '../../../dataset/constant/Element'
import { ElementType } from '../../../dataset/enum/Element'
import { IElement } from '../../../interface/Element'
import { IPasteOption } from '../../../interface/Event'
import {
  getClipboardData,
  getIsClipboardContainFile,
  removeClipboardData
} from '../../../utils/clipboard'
import {
  formatElementContext,
  getElementListByHTML
} from '../../../utils/element'
import { CanvasEvent } from '../CanvasEvent'
import { IOverrideResult } from '../../override/Override'
import { normalizeLineBreak } from '../../../utils'

export function pasteElement(host: CanvasEvent, elementList: IElement[]) {
  console.log('pasteElement1', [...elementList])
  const draw = host.getDraw()
  if (
    draw.isReadonly() ||
    draw.isDisabled() ||
    draw.getControl().getIsDisabledPasteControl()
  ) {
    return
  }
  const rangeManager = draw.getRange()
  const { startIndex } = rangeManager.getRange()
  const originalElementList = draw.getElementList()
  // 全选粘贴无需格式化上下文
  if (~startIndex && !rangeManager.getIsSelectAll()) {
    // 如果是复制到虚拟元素里，则粘贴列表的虚拟元素需扁平化处理，避免产生新的虚拟元素
    const anchorElement = originalElementList[startIndex]

    if (anchorElement?.id || anchorElement?.listId) {
      let start = 0
      while (start < elementList.length) {
        const pasteElement = elementList[start]
        if (anchorElement.id && /^\n/.test(pasteElement.value)) {
          break
        }
        if (VIRTUAL_ELEMENT_TYPE.includes(pasteElement.type!)) {
          elementList.splice(start, 1)
          if (pasteElement.valueList) {
            for (let v = 0; v < pasteElement.valueList.length; v++) {
              const element = pasteElement.valueList[v]
              if (element.value === ZERO || element.value === '\n') {
                continue
              }
              elementList.splice(start, 0, element)
              start++
            }
          }
          start--
        }
        start++
      }
    }
    formatElementContext(originalElementList, elementList, startIndex, {
      isBreakWhenWrap: true,
      editorOptions: draw.getOptions()
    })
  }

  // 使用专门的粘贴函数，避免添加额外的零宽空格和换行符
  console.log('pasteElement', [...elementList])

  draw.insertPastedElementList(elementList)
}

export function pasteHTML(host: CanvasEvent, htmlText: string) {
  const draw = host.getDraw()
  console.log('draw', draw)
  console.log('htmlText', htmlText)
  if (draw.isReadonly() || draw.isDisabled()) return
  const elementList = getElementListByHTML(htmlText, {
    innerWidth: draw.getOriginalInnerWidth()
  })
 console.log('elementList', [...elementList])

  pasteElement(host, elementList)
}

export function pasteImage(host: CanvasEvent, file: File | Blob) {
  const draw = host.getDraw()
  if (draw.isReadonly() || draw.isDisabled()) return
  const rangeManager = draw.getRange()
  const { startIndex } = rangeManager.getRange()
  const elementList = draw.getElementList()

  // 创建FormData对象用于上传
  const formData = new FormData()
  formData.append('file', file)

  // 上传图片到服务器
  fetch(`http://localhost:8000/documents/upload_image/`, {
    method: 'POST',
    body: formData
  })
  .then(response => {
    if (!response.ok) {
      throw new Error(`上传图片失败: ${response.status} ${response.statusText}`)
    }
    return response.json()
  })
  .then(data => {
    // 获取服务器返回的图片URL
    const imageUrl = data.image_url

    // 创建图片对象计算宽高
    const image = new Image()
    image.src = imageUrl

    image.onload = () => {
      const imageElement: IElement = {
        value: imageUrl, // 使用服务器返回的URL替代base64
        type: ElementType.IMAGE,
        width: image.width,
        height: image.height
      }

      if (~startIndex) {
        formatElementContext(elementList, [imageElement], startIndex, {
          editorOptions: draw.getOptions()
        })
      }
      draw.insertElementList([imageElement])
    }

    image.onerror = () => {
      console.error('加载图片失败:', imageUrl)
      // 回退到原始方法，使用base64
      const fileReader = new FileReader()
      fileReader.readAsDataURL(file)
      fileReader.onload = () => {
        const image = new Image()
        const value = fileReader.result as string
        image.src = value
        image.onload = () => {
          const imageElement: IElement = {
            value,
            type: ElementType.IMAGE,
            width: image.width,
            height: image.height
          }
          if (~startIndex) {
            formatElementContext(elementList, [imageElement], startIndex, {
              editorOptions: draw.getOptions()
            })
          }
          draw.insertElementList([imageElement])
        }
      }
    }
  })
  .catch(error => {
    console.error('上传图片失败:', error)
    // 上传失败时回退到原始方法，使用base64
    const fileReader = new FileReader()
    fileReader.readAsDataURL(file)
    fileReader.onload = () => {
      const image = new Image()
      const value = fileReader.result as string
      image.src = value
      image.onload = () => {
        const imageElement: IElement = {
          value,
          type: ElementType.IMAGE,
          width: image.width,
          height: image.height
        }
        if (~startIndex) {
          formatElementContext(elementList, [imageElement], startIndex, {
            editorOptions: draw.getOptions()
          })
        }
        draw.insertElementList([imageElement])
      }
    }
  })
}

export function pasteByEvent(host: CanvasEvent, evt: ClipboardEvent) {
  const draw = host.getDraw()
  if (draw.isReadonly() || draw.isDisabled()) return
  const clipboardData = evt.clipboardData
  if (!clipboardData) return
  // 自定义粘贴事件
  const { paste } = draw.getOverride()
  if (paste) {
    const overrideResult = paste(evt)
    // 默认阻止默认事件
    if ((<IOverrideResult>overrideResult)?.preventDefault !== false) return
  }
  // 优先读取编辑器内部粘贴板数据（粘贴板不包含文件时）
  if (!getIsClipboardContainFile(clipboardData)) {
    const clipboardText = clipboardData.getData('text')
    const editorClipboardData = getClipboardData()
    // 不同系统间默认换行符不同 windows:\r\n mac:\n
    if (
      editorClipboardData &&
      normalizeLineBreak(clipboardText) ===
        normalizeLineBreak(editorClipboardData.text)
    ) {
      pasteElement(host, editorClipboardData.elementList)
      return
    }
  }
  removeClipboardData()
  // 从粘贴板提取数据
  let isHTML = false
  for (let i = 0; i < clipboardData.items.length; i++) {
    const item = clipboardData.items[i]
    if (item.type === 'text/html') {
      isHTML = true
      break
    }
  }
  for (let i = 0; i < clipboardData.items.length; i++) {
    const item = clipboardData.items[i]
    if (item.kind === 'string') {
      if (item.type === 'text/plain' && !isHTML) {
        item.getAsString(plainText => {
          host.input(plainText)
        })
        break
      }
      if (item.type === 'text/html' && isHTML) {
        item.getAsString(htmlText => {
          pasteHTML(host, htmlText)
        })
        break
      }
    } else if (item.kind === 'file') {
      if (item.type.includes('image')) {
        const file = item.getAsFile()
        if (file) {
          pasteImage(host, file)
        }
      }
    }
  }
}

export async function pasteByApi(host: CanvasEvent, options?: IPasteOption) {
  const draw = host.getDraw()
  if (draw.isReadonly() || draw.isDisabled()) return
  // 自定义粘贴事件
  const { paste } = draw.getOverride()
  if (paste) {
    const overrideResult = paste()
    // 默认阻止默认事件
    if ((<IOverrideResult>overrideResult)?.preventDefault !== false) return
  }
  // 优先读取编辑器内部粘贴板数据
  const clipboardText = await navigator.clipboard.readText()
  const editorClipboardData = getClipboardData()
  if (clipboardText === editorClipboardData?.text) {
    pasteElement(host, editorClipboardData.elementList)
    return
  }
  removeClipboardData()
  // 从内存粘贴板获取数据
  if (options?.isPlainText) {
    if (clipboardText) {
      host.input(clipboardText)
    }
  } else {
    const clipboardData = await navigator.clipboard.read()
    let isHTML = false
    for (const item of clipboardData) {
      if (item.types.includes('text/html')) {
        isHTML = true
        break
      }
    }
    for (const item of clipboardData) {
      if (item.types.includes('text/plain') && !isHTML) {
        const textBlob = await item.getType('text/plain')
        const text = await textBlob.text()
        if (text) {
          host.input(text)
        }
      } else if (item.types.includes('text/html') && isHTML) {
        const htmlTextBlob = await item.getType('text/html')
        const htmlText = await htmlTextBlob.text()
        if (htmlText) {
          pasteHTML(host, htmlText)
        }
      } else if (item.types.some(type => type.startsWith('image/'))) {
        const type = item.types.find(type => type.startsWith('image/'))!
        const imageBlob = await item.getType(type)
        pasteImage(host, imageBlob)
      }
    }
  }
}
