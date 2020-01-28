using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDrawIo
{
    public class BiDictionary<TFirst, TSecond>
    {
        private readonly IDictionary<TFirst, TSecond> _firstToSecond;
        private readonly IDictionary<TSecond, TFirst> _secondToFirst;

        public BiDictionary()
          : this(EqualityComparer<TFirst>.Default, EqualityComparer<TSecond>.Default)
        {
        }

        public BiDictionary(IEqualityComparer<TFirst> firstEqualityComparer, IEqualityComparer<TSecond> secondEqualityComparer)
        {
            _firstToSecond = new Dictionary<TFirst, TSecond>(firstEqualityComparer);
            _secondToFirst = new Dictionary<TSecond, TFirst>(secondEqualityComparer);
        }

        public void Add(TFirst first, TSecond second)
        {
            TFirst existingFirst;
            TSecond existingSecond;

            if (_firstToSecond.TryGetValue(first, out existingSecond))
            {
                if (!existingSecond.Equals(second))
                    throw new ArgumentException($"Duplicate item already exists for '{first}'.");
            }

            if (_secondToFirst.TryGetValue(second, out existingFirst))
            {
                if (!existingFirst.Equals(first))
                    throw new ArgumentException($"Duplicate item already exists for '{second}'.");
            }

            _firstToSecond.Add(first, second);
            _secondToFirst.Add(second, first);
        }

        public void AddOrUpdate(TFirst first, TSecond second)
        {
            TryRemoveByFirst(first, out _);
            Add(first, second);
        }

        public bool TryGetByFirst(TFirst first, out TSecond second)
        {
            return _firstToSecond.TryGetValue(first, out second);
        }

        public bool TryGetBySecond(TSecond second, out TFirst first)
        {
            return _secondToFirst.TryGetValue(second, out first);
        }

        public bool TryRemoveByFirst(TFirst first, out TSecond second)
        {
            if (_firstToSecond.TryGetValue(first, out second))
            {
                _firstToSecond.Remove(first);
                _secondToFirst.Remove(second);

                return true;
            }

            return false;
        }

        public bool TryRemoveBySecond(TSecond second, out TFirst first)
        {
            if (_secondToFirst.TryGetValue(second, out first))
            {
                _secondToFirst.Remove(second);
                _firstToSecond.Remove(first);
                
                return true;
            }

            return false;
        }

        public ICollection<TFirst> FirstKeys() => _firstToSecond.Keys;

        public ICollection<TSecond> SecondKeys => _secondToFirst.Keys;
    }

}
