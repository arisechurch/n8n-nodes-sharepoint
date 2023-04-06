import type { OptionsWithUri } from "request"
import type {
  IDataObject,
  IExecuteFunctions,
  IExecuteSingleFunctions,
  ILoadOptionsFunctions,
  INodeExecutionData,
  INodePropertyOptions,
  JsonObject,
} from "n8n-workflow"
import { NodeApiError } from "n8n-workflow"
import type { Effect } from "@effect/io/Effect"
import { Tag } from "@effect/data/Context"
import type { Option } from "@effect/data/Option"
import { TaggedClass } from "@effect/data/Data"
import { identity } from "@effect/data/Function"

export const N8NId = Symbol.for("N8N")
export interface N8N {
  [N8NId]: typeof N8NId
}
export const N8N = Tag<
  N8N,
  IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions
>()

export function execute<E, A>(
  effect: Effect<N8N, E, A>,
): (
  this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions,
) => Promise<A> {
  return function () {
    return effect.provideService(N8N, this).runPromise
  }
}

export function microsoftApiRequest<A = any>(
  method: string,
  resource: string,
  {
    qs = {},
    body,
    uri,
    headers = {},
  }: {
    body?: any
    qs?: IDataObject
    uri?: string
    headers?: IDataObject
  } = {},
): Effect<N8N, NodeApiError, A> {
  const options: OptionsWithUri = {
    headers: {
      ...headers,
      "Content-Type": "application/json",
    },
    method,
    body,
    qs,
    uri: uri || `https://graph.microsoft.com/v1.0${resource}`,
    json: true,
  }

  return Do($ => {
    const n8n = $(N8N)

    return $(
      Effect.tryCatchPromise(
        () =>
          n8n.helpers.requestOAuth2.call(
            n8n,
            "microsoftSharepointApi",
            options,
          ),
        error =>
          new NodeApiError(n8n.getNode(), error as JsonObject, error as any),
      ),
    )
  })
}

export interface Response<A> {
  readonly value: A
}

// ==== sites
export interface SharepointSite {
  readonly id: string
  readonly displayName: string
  readonly name: string
}

export const sites = microsoftApiRequest<Response<SharepointSite[]>>(
  "GET",
  "/sites",
  {
    qs: { search: "" },
  },
)
export const siteOptions = sites.map(({ value }) =>
  value.map(
    (_): INodePropertyOptions => ({
      name: _.displayName,
      value: _.id.split(",")[1],
    }),
  ),
)

export class NoSuchParam extends TaggedClass("NoSuchParam")<{
  readonly name: string
}> {
  readonly message = `param not set ${this.name}`
}

export const getOptionalParam = <A = string>(
  name: string,
  i = 0,
): Effect<N8N, never, Option<A>> =>
  Do($ => {
    const n8n = $(N8N)
    const getParam = Option.liftThrowable(n8n.getNodeParameter.bind(n8n))
    return (getParam(name, i) as Option<A>).filter(_ => !!_)
  })

export const getParam = <A = string>(name: string, i = 0) =>
  getOptionalParam<A>(name, i)
    .flatMap(identity)
    .mapError(_ => new NoSuchParam({ name }))

const siteId = getParam("site")

// ==== lists
export interface SharepointList {
  readonly id: string
  readonly displayName: string
  readonly name: string
  readonly list: {
    readonly template: string
  }
}

export const getLists = Do($ => {
  const site = $(siteId)
  return $(
    microsoftApiRequest<Response<SharepointList[]>>(
      "GET",
      `/sites/${site}/lists`,
    ).map(_ => _.value.filter(_ => _.list.template === "genericList")),
  )
})

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
    microsoftApiRequest(method, `/sites/${siteId}/lists/${listId}${path}`, {
      body: body.getOrUndefined,
    })
      .map(_ => _.value ?? _)
      .map(_ =>
        Array.isArray(_) ? _.map(wrapExecutionData) : [wrapExecutionData(_)],
      ),
  )
})

export function wrapExecutionData<A extends IDataObject>(
  json: A,
): INodeExecutionData {
  return { json }
}
