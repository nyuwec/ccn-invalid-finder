export class MMap<K, V> extends Map<K, V> {
  getOrElse(key: K, defaultValue: V): V {
    return this.has(key) ? this.get(key) as V : defaultValue
  }
}
