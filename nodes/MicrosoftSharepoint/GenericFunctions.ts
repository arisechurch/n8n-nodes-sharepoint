import type { INodePropertyOptions } from "n8n-workflow"
import {
  N8N,
  Response,
  apiRequest,
  getOptionalParam,
  getParam,
  wrapData,
  wrapJson,
} from "../common"
import type { Matcher } from "@effect/match"
import type { Effect } from "@effect/io/Effect"

// ==== sites
export interface SharepointSite {
  readonly id: string
  readonly displayName: string
  readonly name: string
}

export const sites = apiRequest<Response<SharepointSite[]>>("GET", "/sites", {
  qs: { search: "" },
})

export const siteOptions = sites.map(({ value }) =>
  value.map(
    (_): INodePropertyOptions => ({
      name: _.displayName,
      value: _.id.split(",")[1],
    }),
  ),
)

// ==== lists
export interface SharepointList {
  readonly id: string
  readonly displayName: string
  readonly name: string
  readonly list: {
    readonly hidden: boolean
    readonly template: string
  }
}

export const getLists = getParam("site")
  .flatMap(siteId =>
    apiRequest<Response<SharepointList[]>>("GET", `/sites/${siteId}/lists`),
  )
  .map(_ => _.value.filter(({ list }) => !list.hidden))

export const listOptions = getLists.map(_ =>
  _.map(
    (_): INodePropertyOptions => ({
      name: _.displayName,
      value: _.id,
    }),
  ),
)

// ==== files
export type SharepointItem = SharepointFile | SharepointFolder

export interface SharepointFile {
  readonly id: string
  readonly name: string
  readonly file: {
    readonly mimeType: string
  }
}

export interface SharepointFolder {
  readonly id: string
  readonly name: string
  readonly folder: {
    readonly childCount: number
  }
}

const isFile = (item: SharepointItem): item is SharepointFile => "file" in item
const isFolder = (item: SharepointItem): item is SharepointFolder =>
  "folder" in item

const folderPath = getParam("folder").map(_ => _.replace(/^\/+/, ""))

const getItem = Do($ => {
  const siteId = $(getParam("site"))
  const path = $(folderPath)
  return $(
    apiRequest<SharepointItem>(
      "GET",
      path.length > 0
        ? `/sites/${siteId}/drive/root:/${path}`
        : `/sites/${siteId}/drive/root`,
    ),
  )
})

const getItems = Do($ => {
  const siteId = $(getParam("site"))
  const path = $(folderPath)
  const children = $(
    apiRequest<Response<SharepointItem[]>>(
      "GET",
      path.length > 0
        ? `/sites/${siteId}/drive/root:/${path}:/children`
        : `/sites/${siteId}/drive/root/children`,
    ),
  )
  return children.value
})

const getFiles = getItems.map(_ => _.filter(isFile))
const getFolders = Effect.all([getItem, getItems]).map(([self, children]) =>
  [self, ...children].filter(isFolder),
)

export const fileOptions = getFiles.map(_ =>
  _.map(file => ({
    name: file.name,
    value: file.id,
  })),
)

export const folderOptions = getFolders.map(_ =>
  _.map(file => ({
    name: file.name,
    value: file.id,
  })),
)

// === execution
const call = <A = any>(url: string) =>
  Do($ => {
    const method = $(getParam("method"))
    const body = $(getOptionalParam("body"))

    return $(
      apiRequest<A>(method, url, {
        body: body.getOrUndefined,
      }),
    )
  })
const callBinary = (url: string) =>
  apiRequest<Buffer>("GET", url, { encoding: null })

const getOp = Matcher.type<string>()
  .when("lists", () =>
    Do($ => {
      const siteId = $(getParam("site"))
      const listId = $(getParam("list"))
      const path = $(getParam("path"))
      const response = $(call(`/sites/${siteId}/lists/${listId}${path}`))
      return [wrapJson(response)]
    }),
  )
  .when("files", () =>
    Effect.gen(function* ($) {
      const n8n = yield* $(N8N)
      const siteId = yield* $(getParam("site"))
      const fileId = yield* $(getParam("fileId"))
      const path = yield* $(getParam("path"))

      if (path.startsWith("/content")) {
        const details = yield* $(
          call<SharepointFile>(`/sites/${siteId}/drive/items/${fileId}`),
        )
        const data = yield* $(
          callBinary(`/sites/${siteId}/drive/items/${fileId}/content`).flatMap(
            _ =>
              Effect.promise(() =>
                n8n.helpers.prepareBinaryData(
                  _,
                  details.name,
                  details.file.mimeType,
                ),
              ),
          ),
        )

        return [
          wrapData({
            json: details as any,
            binary: { data },
          }),
        ]
      }

      const response = yield* $(
        call(`/sites/${siteId}/drive/items/${fileId}${path}`),
      )
      return [wrapJson(response)]
    }),
  )
  .when("folders", () =>
    Do($ => {
      const siteId = $(getParam("site"))
      const folderId = $(getParam("folderId"))
      const path = $(getParam("path"))
      const response = $(
        call(`/sites/${siteId}/drive/items/${folderId}${path}`),
      )
      return [wrapJson(response)]
    }),
  )
  .orElse(() => Effect.die(new Error("Invalid resource")))

export const run = Do($ => {
  const resource = $(getParam("resource"))
  return $(getOp(resource))
})
