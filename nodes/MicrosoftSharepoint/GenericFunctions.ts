import type { INodePropertyOptions } from "n8n-workflow"
import {
  Response,
  apiRequest,
  getOptionalParam,
  getParam,
  wrapJson,
} from "../common"

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

// === execution
export const run = Do($ => {
  const siteId = $(getParam("site"))
  const listId = $(getParam("list"))
  const method = $(getParam("method"))
  const path = $(getParam("path"))
  const body = $(getOptionalParam("body"))

  return $(
    apiRequest(method, `/sites/${siteId}/lists/${listId}${path}`, {
      body: body.getOrUndefined,
    }).map(wrapJson),
  )
})
