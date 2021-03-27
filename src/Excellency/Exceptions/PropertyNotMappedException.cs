using System;

namespace Excellency.Exceptions
{
    public class PropertyNotMappedException :Exception
    {
        public PropertyNotMappedException()
        {
        }

        public PropertyNotMappedException(string message)
            : base(message)
        {
        }

        public PropertyNotMappedException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
